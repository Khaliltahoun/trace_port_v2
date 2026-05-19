(function () {
  "use strict";

  const DATA = window.PFE_DATA;
  const EVENTS_KEY = "trace-port-digital-events-v1";
  const TRAINS_KEY = "trace-port-digital-trains-v1";
  const SHIPS_KEY = "trace-port-digital-ships-v1";
  const VALIDATIONS_KEY = "trace-port-validations-v1";
  const LOGS_KEY = "trace-port-logs-v1";
  const PROFILE_KEY = "trace-port-active-profile-v1";

  const PROFILES = [
    { id: "agent", name: "Ahmed Benali", role: "Agent de quart", initials: "AB", scope: "Saisie des arrêts et suivi des incidents terrain", defaultView: "entry", color: "#1d4ed8" },
    { id: "chef", name: "Youssef El Amrani", role: "Chef d'équipe", initials: "YE", scope: "Validation des arrêts et coordination quart", defaultView: "validation", color: "#0f766e" },
    { id: "manutention", name: "El Houdzi Israa", role: "Responsable manutention", initials: "EH", scope: "Pilotage temps réel de l'exploitation manutention", defaultView: "dashboard", color: "#0b315f" },
    { id: "performance", name: "Khadija Saidi", role: "Responsable performance", initials: "KS", scope: "Analyse KPI, Pareto, performance circuits, alertes", defaultView: "performance", color: "#7c5ce6" },
    { id: "admin", name: "Omar Lahbabi", role: "Administrateur", initials: "OL", scope: "Référentiels, RBAC, configuration et audit", defaultView: "logs", color: "#dc2626" }
  ];

  const ROLE_PERMISSIONS = {
    "Agent de quart": { canValidate: false, canCreate: true, focus: "execution" },
    "Chef d'équipe": { canValidate: true, canCreate: true, focus: "validation" },
    "Responsable manutention": { canValidate: true, canCreate: false, focus: "operations" },
    "Responsable performance": { canValidate: false, canCreate: false, focus: "performance" },
    "Administrateur": { canValidate: true, canCreate: true, focus: "governance" }
  };

  const COST_PER_STOP_HOUR_EUR = 3850;
  const TRS_TARGET = 0.85;
  const TRS_MAINT_TARGET = 0.80;

  /* ============================================================
   * OFFICIAL CALCULATION ENGINE
   * Reproduces the perimeter of the Excel "Synthèses" sheet so
   * TRACE-PORT totals are perfectly aligned with the legacy
   * monthly report sent to direction.
   * ============================================================ */

  // 32 canonical families matching the Synthèses!row10 headers (case + accents)
  const OFFICIAL_FAMILIES = [
    "exploitation", "électrique", "instrumentation", "mécanique", "bande",
    "manque navire", "attente accostage", "attente préparation", "structure navire",
    "Mouvement portique", "Changement de cale", "Mouvement même cale", "Passage marée",
    "intempéries", "Assiette", "Balance", "LTE", "Déhalage", "Déballastage",
    "Correction gîte", "Navire haut", "Sondage des ballasts", "Arrêt par le bord",
    "Arrêt par le quai", "Forces majeurs", "stock", "arrêts planifiés",
    "manque produit", "Alimentation du train", "Dégagement stériles",
    "Permutation circuit", "autres externes"
  ];

  // Raw label → official family alias map (uses normalized lookup).
  // These variants were identified by reconciliation against the actual Bilan dataset.
  const FAMILY_ALIASES = {
    "electrique": "électrique",
    "éléctrique": "électrique",
    "electrique+mecanique": "électrique",
    "attente acostage": "attente accostage",
    "passage pluie": "intempéries",
    "intemperies": "intempéries",
    "travaux planifier": "arrêts planifiés",
    "travaux planifies": "arrêts planifiés",
    "arrêts planifiées": "arrêts planifiés",
    "arrets planifies": "arrêts planifiés",
    "mecanique": "mécanique",
    "deballastage": "Déballastage",
    "dehalage": "Déhalage",
    "balance": "Balance",
    "lte": "LTE",
    "passage maree": "Passage marée",
    "marée": "Passage marée",
    "assiette": "Assiette",
    "correction gite": "Correction gîte",
    "correction gîte": "Correction gîte",
    "navire haut": "Navire haut",
    "sondage": "Sondage des ballasts",
    "sondage des ballasts": "Sondage des ballasts",
    "arret par le bord": "Arrêt par le bord",
    "arret par le quai": "Arrêt par le quai",
    "forces majeurs": "Forces majeurs",
    "force majeure": "Forces majeurs",
    "alimentation du train": "Alimentation du train",
    "alimentation train": "Alimentation du train",
    "degagement steriles": "Dégagement stériles",
    "permutation circuit": "Permutation circuit",
    "autres externes": "autres externes",
    "mouvement portique": "Mouvement portique",
    "changement de cale": "Changement de cale",
    "mouvement meme cale": "Mouvement même cale",
    "mouvement même cale": "Mouvement même cale",
    "manque navire": "manque navire",
    "manque produit": "manque produit",
    "structure navire": "structure navire",
    "attente accostage": "attente accostage",
    "attente préparation": "attente préparation",
    "attente preparation": "attente préparation",
    "stock": "stock",
    "exploitation": "exploitation",
    "instrumentation": "instrumentation",
    "bande": "bande"
  };

  // Excel Synthèses!Row15 reference baseline — per-family CHARGING totals
  // extracted from the same Janvier 2026 source workbook that built the
  // current TRACE-PORT dataset. Stored as the immutable comparison reference
  // so manager-level reconciliation matches Excel exactly at TRACE-PORT
  // startup, and any divergence visible afterwards is caused by new manual
  // entries / corrections inside TRACE-PORT.
  const EXCEL_REFERENCE_JAN_2026 = {
    "exploitation": 1.17,
    "électrique": 6.08,
    "instrumentation": 0,
    "mécanique": 6.75,
    "bande": 2.17,
    "manque navire": 768.00,
    "attente accostage": 307.17,
    "attente préparation": 101.33,
    "structure navire": 193.25,
    "Mouvement portique": 0,
    "Changement de cale": 11.75,
    "Mouvement même cale": 11.92,
    "Passage marée": 2.67,
    "intempéries": 967.50,
    "Assiette": 4.08,
    "Balance": 31.17,
    "LTE": 18.08,
    "Déhalage": 0,
    "Déballastage": 9.67,
    "Correction gîte": 0,
    "Navire haut": 0,
    "Sondage des ballasts": 0,
    "Arrêt par le bord": 2.67,
    "Arrêt par le quai": 3.83,
    "Forces majeurs": 0,
    "stock": 78.92,
    "arrêts planifiés": 170.58,
    "manque produit": 40.00,
    "Alimentation du train": 6.00,
    "Dégagement stériles": 4.33,
    "Permutation circuit": 7.08,
    "autres externes": 0
  };

  // KPI definitions — single source of truth, used for tooltips & docs
  const KPI_DEFINITIONS = {
    totalStopHours: {
      label: "Temps d'arrêt total (officiel)",
      formula: "Σ durationHours pour événements (S/E ∈ CA30/CB30/CC30/CD30) ET famille ∈ taxonomie officielle",
      sheet: "Synthèses!B15"
    },
    trsExploitation: {
      label: "TRS Exploitation",
      formula: "(heures dispo − arrêts exploitation) / heures dispo",
      sheet: "Synthèses!AP20"
    },
    trsMaintenance: {
      label: "TRS Maintenance",
      formula: "(heures dispo − arrêts maintenance: électrique+instrumentation+mécanique+bande) / heures dispo",
      sheet: "Synthèses!AQ20"
    },
    trsGlobal: {
      label: "TRS Global",
      formula: "(heures dispo − arrêts exploitation − arrêts maintenance) / heures dispo",
      sheet: "Synthèses!AR20"
    },
    cadenceTph: {
      label: "Cadence horaire (t/h)",
      formula: "tonnage pesage / heures de marche",
      sheet: "Synthèses!H20"
    }
  };

  const CIRCUITS = [
    { key: "CA30", role: "Circuit chargement 1", color: "#1aa872" },
    { key: "CB30", role: "Circuit chargement 2", color: "#2563eb" },
    { key: "CC30", role: "Circuit chargement 3", color: "#7558e0" },
    { key: "CD30", role: "Circuit chargement 4", color: "#ea7621" }
  ];
  const SILOS = [
    { key: "DA", label: "Silo DA", role: "Déchargement train · poste DA10" },
    { key: "DB", label: "Silo DB", role: "Déchargement train · poste DB10" }
  ];
  const QUALITIES = ["K01", "K02", "K03", "K08", "K09s", "K10", "K12", "K20", "K62"];

  const STOP_TEMPLATES = [
    { family: "Balance", desc: "Attente balance — synchronisation pesage", typical: 60, scope: "ship" },
    { family: "intempéries", desc: "Passage pluie — protection produit", typical: 290, scope: "global" },
    { family: "attente accostage", desc: "Navire en attente — accostage non débuté", typical: 405, scope: "global" },
    { family: "attente préparation", desc: "Préparation cale en cours", typical: 240, scope: "ship" },
    { family: "Changement de cale", desc: "Mouvement portique vers nouvelle cale", typical: 45, scope: "ship" },
    { family: "Mouvement même cale", desc: "Mouvement portique intra-cale", typical: 35, scope: "ship" },
    { family: "Passage marée", desc: "Stop sécurité passage de marée", typical: 40, scope: "global" },
    { family: "LTE", desc: "Limite Tonnage Embarqué — vérification douane", typical: 60, scope: "ship" },
    { family: "Assiette", desc: "Correction d'assiette navire", typical: 75, scope: "ship" },
    { family: "Déhalage", desc: "Déhalage navire le long du quai", typical: 30, scope: "ship" },
    { family: "manque navire", desc: "Aucun navire à charger", typical: 240, scope: "global" },
    { family: "stock", desc: "Manque produit en silo", typical: 180, scope: "stock" },
    { family: "exploitation", desc: "Débit faible / surcharge ligne", typical: 30, scope: "circuit" },
    { family: "mécanique", desc: "Intervention mécanique sur équipement", typical: 90, scope: "circuit" },
    { family: "électrique", desc: "Défaillance électrique", typical: 60, scope: "circuit" },
    { family: "bande", desc: "Travaux bande transporteuse", typical: 120, scope: "circuit" },
    { family: "Alimentation du train", desc: "Alimentation rame train", typical: 90, scope: "stock" }
  ];

  const VIEW_TITLES = {
    dashboard: "Centre de commandement",
    entry: "Justifier un arrêt",
    myStops: "Mes arrêts",
    currentStops: "Arrêts en cours",
    validation: "File de validation",
    pareto: "Analyse Pareto",
    performance: "Performance circuits",
    reporting: "Reporting & exports",
    monthlySynth: "Synthèse mensuelle",
    trains: "Trains & déchargement",
    stocks: "Stocks & silos",
    ships: "Navires & chargement",
    references: "Référentiels",
    equipments: "Équipements & circuits",
    stopNatures: "Natures d'arrêt",
    users: "Utilisateurs & rôles",
    settings: "Paramètres plateforme",
    logs: "Logs & traçabilité",
    reconciliation: "Réconciliation Excel",
    stopDetail: "Fiche incident",
    daily: "Synthèse journalière",
    monthly: "Synthèse mensuelle",
    events: "Journal des arrêts",
    tonnage: "Tonnage",
    flow: "Trains & navires",
    formulas: "Formules & requêtes",
    dmaic: "Besoins PFE"
  };

  const VIEW_ALIASES = {
    kpiDashboard: "dashboard",
    indicators: "dashboard",
    dailyReports: "reporting",
    monthlyReports: "reporting",
    exportData: "reporting",
    flow: "trains",
    monthly: "monthlySynth"
  };

  const VIEW_META = {
    dashboard: ["Pilotage", "Centre de commandement temps réel", "Surveiller circuits, KPI et incidents critiques", "currentStops"],
    entry: ["Workflow arrêts", "Justification horaire d'arrêt", "Saisir un arrêt et propager sur les circuits affectés", "myStops"],
    myStops: ["Workflow arrêts", "Suivi de mes arrêts", "Consulter le statut et corriger les arrêts rejetés", "validation"],
    currentStops: ["Workflow arrêts", "Arrêts actifs et critiques", "Prioriser les arrêts longs et déclencher la validation", "validation"],
    validation: ["Workflow arrêts", "Contrôle Chef d'équipe", "Valider ou rejeter pour fiabiliser les KPI", "dashboard"],
    stopDetail: ["Workflow arrêts", "Fiche complète d'incident", "Consulter l'historique puis valider ou corriger", "validation"],
    performance: ["Pilotage", "Performance des circuits manutention", "Identifier les circuits sous objectif", "pareto"],
    pareto: ["Pilotage", "Diagnostic des causes racines", "Prioriser les leviers d'amélioration", "reporting"],
    reporting: ["Reporting", "Diffusion & extractions contrôlées", "Générer la liasse opérationnelle attendue par les services", "monthlySynth"],
    monthlySynth: ["Reporting", "Synthèse mensuelle officielle", "Consolider le bilan livré à la direction en fin de mois", "reporting"],
    trains: ["Chaîne logistique", "Trains & déchargement", "Suivre l'arrivée des trains et l'alimentation des silos", "stocks"],
    stocks: ["Chaîne logistique", "Stocks & silos", "Visualiser la circulation produit entre silos DA/DB", "ships"],
    ships: ["Chaîne logistique", "Navires & chargement", "Piloter le chargement, les peseuses et l'écart connaissement", "entry"],
    references: ["Administration", "Référentiels industriels", "Maintenir équipements, circuits et natures d'arrêt", "users"],
    equipments: ["Administration", "Référentiel équipements & circuits", "Maintenir la cartographie industrielle", "stopNatures"],
    stopNatures: ["Administration", "Référentiel natures d'arrêt", "Normaliser les familles de causes", "users"],
    users: ["Administration", "Utilisateurs & RBAC", "Gérer les rôles et permissions", "logs"],
    logs: ["Administration", "Logs & traçabilité", "Contrôler les actions sensibles", "settings"],
    settings: ["Administration", "Configuration plateforme", "Piloter les paramètres système", "dashboard"],
    reconciliation: ["Administration", "Réconciliation Excel ↔ TRACE-PORT", "Auditer l'alignement entre la synthèse officielle et le calcul live", "logs"]
  };

  const VIEWS_WITH_FILTERS = new Set(["dashboard", "myStops", "currentStops", "validation", "pareto", "performance", "events"]);
  const VIEWS_WITH_EXPORTS = new Set(["dashboard", "myStops", "currentStops", "validation", "pareto", "performance", "events", "reporting", "monthlySynth", "trains", "ships"]);
  const VIEWS_WITH_PERIOD = new Set(["dashboard", "reporting", "pareto", "performance", "monthlySynth", "trains", "ships", "stocks"]);
  const CHARGING_SECTIONS = ["CA30", "CB30", "CC30", "CD30"];
  const DISCHARGE_SECTIONS = ["DA10", "DB10"];
  const USERS = [
    { firstName: "Ahmed", lastName: "Benali", role: "Agent de quart", service: "Exploitation", status: "Actif" },
    { firstName: "Youssef", lastName: "El Amrani", role: "Chef d'équipe", service: "Exploitation", status: "Actif" },
    { firstName: "Khadija", lastName: "Saidi", role: "Superviseur", service: "Maintenance", status: "Actif" },
    { firstName: "Omar", lastName: "Lahbabi", role: "Administrateur", service: "Système", status: "Actif" },
    { firstName: "Israa", lastName: "El Houdzi", role: "Responsable", service: "DLP Casablanca", status: "Actif" }
  ];
  const ROLES = [
    { name: "Agent", permissions: "Créer un arrêt, consulter ses arrêts" },
    { name: "Chef d'équipe", permissions: "Valider, rejeter, corriger les arrêts" },
    { name: "Superviseur", permissions: "Piloter KPI, rapports et alertes" },
    { name: "Responsable", permissions: "Consulter tableaux de bord et rapports consolidés" },
    { name: "Administrateur", permissions: "Gérer référentiels, utilisateurs et paramètres" }
  ];
  const MAINTENANCE_FAMILIES = ["électrique", "instrumentation", "mécanique", "bande"];
  const EXTERNAL_FAMILIES = [
    "manque navire",
    "attente accostage",
    "attente préparation",
    "structure navire",
    "passage marée",
    "intempéries",
    "assiette",
    "balance",
    "lte",
    "déhalage",
    "déballastage",
    "correction gîte",
    "navire haut",
    "sondage des ballasts",
    "arrêt par le bord",
    "arrêt par le quai",
    "forces majeurs",
    "arrêts planifiés",
    "manque produit",
    "autres externes"
  ];

  const FORMULA_BLUEPRINTS = [
    {
      module: "Durée arrêt",
      excel: "Bilan!F = fin - début",
      digital: "durationHours = (end - start) * 24"
    },
    {
      module: "Synthèse par famille",
      excel: "SUMIFS(Bilan!F:F, Bilan!A:A, S/E, Bilan!C:C, famille)",
      digital: "sumIfs(events, { sectionKey, family })"
    },
    {
      module: "Total arrêts",
      excel: "Synthèses!B = SUM(familles D:AI)",
      digital: "totalStopHours = sum(durationHours)"
    },
    {
      module: "Maintenance",
      excel: "Synthèses!AL = SUM(électrique: bande)",
      digital: "maintenanceHours = électrique + instrumentation + mécanique + bande"
    },
    {
      module: "TRS exploitation",
      excel: "AP = (durée affectation - arrêt exploitation) / durée affectation",
      digital: "trsExploitation = (availableHours - exploitationHours) / availableHours"
    },
    {
      module: "TRS maintenance",
      excel: "AQ = (durée affectation - arrêts maintenance) / durée affectation",
      digital: "trsMaintenance = (availableHours - maintenanceHours) / availableHours"
    },
    {
      module: "TRS global",
      excel: "AR = (durée affectation - maintenance - exploitation) / durée affectation",
      digital: "trsGlobal = (availableHours - maintenanceHours - exploitationHours) / availableHours"
    },
    {
      module: "Cadence horaire",
      excel: "Débit = tonnage / heures de marche",
      digital: "cadenceTph = tonnage / runningHours"
    },
    {
      module: "Tonnage journalier",
      excel: "Tonnage!M = SUM(C:L), Tonnage!Z = SUM(P:Y)",
      digital: "pesageTotal = sum(qualities), draftTotal = sum(qualities)"
    },
    {
      module: "Ecart navire",
      excel: "Navire!O = (connaissement - bascule) / connaissement",
      digital: "gapRatio = (billOfLading - scaleTotal) / billOfLading"
    }
  ];

  const PAGE_SIZE = 20;
  const AUTO_REFRESH_MS = 30000;
  const MONTH_NAMES = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];

  function datasetMonthKey() {
    // Derive default selected month from the dataset's most populated month
    const counts = new Map();
    (DATA?.events || []).forEach((e) => {
      const ts = e.start || e.end;
      if (!ts) return;
      const k = String(ts).slice(0, 7);
      counts.set(k, (counts.get(k) || 0) + 1);
    });
    let best = "2026-01";
    let max = 0;
    for (const [k, v] of counts.entries()) {
      if (v > max) { max = v; best = k; }
    }
    return best;
  }

  const state = {
    view: "dashboard",
    profile: loadProfile(),
    authenticated: loadAuthFlag(),
    period: "month",
    selectedMonth: localStorage.getItem("trace-port-selected-month") || datasetMonthKey(),
    analysisDate: localStorage.getItem("trace-port-analysis-date") || defaultAnalysisDate(),
    calcMode: localStorage.getItem("trace-port-calc-mode") || "official",
    customFrom: "",
    customTo: "",
    filters: {
      section: "all",
      family: "all",
      quality: "all",
      search: ""
    },
    dailyDate: null,
    selectedEventId: null,
    formulaSearch: "",
    formulaSheet: "all",
    reportingPeriod: "daily",
    lastSync: new Date(),
    autoRefreshTimer: null,
    pagination: {
      myStops: 1,
      currentStops: 1,
      validation: 1,
      events: 1,
      logs: 1,
      equipments: 1,
      stopNatures: 1,
      formulas: 1
    }
  };

  function loadAuthFlag() {
    try { return localStorage.getItem("trace-port-auth-v1") === "true"; } catch { return false; }
  }

  function saveAuthFlag(value) {
    try { localStorage.setItem("trace-port-auth-v1", value ? "true" : "false"); } catch {}
  }

  function startOfDay(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
  function endOfDay(d) { const x = new Date(d); x.setHours(23, 59, 59, 999); return x; }

  function defaultAnalysisDate() {
    const monthKey = localStorage.getItem("trace-port-selected-month") || datasetMonthKey();
    return monthEndDate(monthKey);
  }

  function monthEndDate(monthKey) {
    const [year, month] = monthKey.split("-").map(Number);
    const end = new Date(year || 2026, month || 1, 0);
    return dateKey(end);
  }

  function analysisAnchorDate() {
    const value = state.analysisDate || defaultAnalysisDate();
    const parsed = new Date(`${value}T12:00:00`);
    return Number.isFinite(parsed.getTime()) ? parsed : new Date();
  }

  function periodRange() {
    const anchor = analysisAnchorDate();
    const range = { from: null, to: null, label: "" };
    if (state.period === "today") {
      const today = startOfDay(anchor);
      range.from = today; range.to = endOfDay(anchor);
      range.label = `Aujourd'hui · ${today.toLocaleDateString("fr-FR")}`;
    } else if (state.period === "7d") {
      const end = endOfDay(anchor);
      const start = startOfDay(new Date(anchor.getTime() - 6 * 86400000));
      range.from = start; range.to = end;
      range.label = "7 derniers jours";
    } else if (state.period === "mtd") {
      const start = new Date(anchor.getFullYear(), anchor.getMonth(), 1, 0, 0, 0);
      range.from = start; range.to = endOfDay(anchor);
      range.label = `Mois en cours · ${MONTH_NAMES[start.getMonth()]} ${start.getFullYear()}`;
    } else if (state.period === "custom" && state.customFrom && state.customTo) {
      range.from = startOfDay(new Date(`${state.customFrom}T00:00:00`));
      range.to = endOfDay(new Date(`${state.customTo}T00:00:00`));
      range.label = `Du ${range.from.toLocaleDateString("fr-FR")} au ${range.to.toLocaleDateString("fr-FR")}`;
    } else {
      const [y, m] = (state.selectedMonth || datasetMonthKey()).split("-").map(Number);
      const monthIdx = (m || 1) - 1;
      const start = new Date(y, monthIdx, 1, 0, 0, 0);
      const end = new Date(y, monthIdx + 1, 0, 23, 59, 59);
      range.from = start; range.to = end;
      range.label = `${MONTH_NAMES[monthIdx]} ${y} complet`;
    }
    return range;
  }

  function periodDays() {
    const r = periodRange();
    if (!r.from || !r.to) return 31;
    const diff = (r.to.getTime() - r.from.getTime()) / 86400000;
    return Math.max(1, Math.ceil(diff));
  }

  function periodElapsedFraction() {
    if (state.period !== "mtd") return 1;
    const r = periodRange();
    const monthStart = new Date(r.from);
    const monthEnd = new Date(r.from.getFullYear(), r.from.getMonth() + 1, 0, 23, 59, 59);
    const totalMs = monthEnd.getTime() - monthStart.getTime();
    const elapsedMs = Math.max(0, Math.min(totalMs, r.to.getTime() - monthStart.getTime()));
    return totalMs > 0 ? elapsedMs / totalMs : 1;
  }

  function eventInPeriod(event) {
    const r = periodRange();
    if (!r.from || !r.to) return true;
    const ts = new Date(event.start || event.end || event.day || 0).getTime();
    if (!Number.isFinite(ts) || ts === 0) return false;
    return ts >= r.from.getTime() && ts <= r.to.getTime();
  }

  function hasDataForPeriod() {
    return getFilteredEvents().length > 0
      || getAllShips().some((s) => s.name && eventInPeriod({ start: s.start }))
      || getAllTrains().some((t) => eventInPeriod({ start: t.day }));
  }

  function loadProfile() {
    try {
      const stored = localStorage.getItem(PROFILE_KEY);
      if (stored) {
        const found = PROFILES.find((p) => p.id === stored);
        if (found) return found;
      }
    } catch {}
    return PROFILES[0];
  }

  function saveProfile(id) {
    try { localStorage.setItem(PROFILE_KEY, id); } catch {}
  }

  function currentUser() {
    return state.profile;
  }

  function currentRolePermissions() {
    return ROLE_PERMISSIONS[state.profile.role] || {};
  }

  const WORKFLOW_STEPS = [
    { key: "Saisie", target: "entry", views: ["entry", "myStops"] },
    { key: "Validation", target: "validation", views: ["validation", "stopDetail"] },
    { key: "Pilotage", target: "dashboard", views: ["dashboard", "performance"] },
    { key: "Analyse", target: "pareto", views: ["pareto"] },
    { key: "Reporting", target: "reporting", views: ["reporting"] },
    { key: "Gouvernance", target: "logs", views: ["equipments", "stopNatures", "users", "logs", "settings"] }
  ];

  const NAV_ICONS = {
    command: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 13h6V3H3v10Zm0 8h6v-6H3v6Zm8 0h10V11H11v10Zm0-18v6h10V3H11Z" fill="currentColor"/></svg>',
    gauge: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4a8 8 0 0 0-8 8h2a6 6 0 0 1 12 0h2a8 8 0 0 0-8-8Zm0 3v5l3.5 3.5 1.4-1.4L13 11.2V7h-1Z" fill="currentColor"/></svg>',
    pareto: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20V4h2v16H4Zm4 0V10h2v10H8Zm4 0V13h2v7h-2Zm4 0V8h2v12h-2Zm4 0V5h0V4h-2v16h2Z" fill="currentColor"/></svg>',
    add: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M11 5v6H5v2h6v6h2v-6h6v-2h-6V5h-2Z" fill="currentColor"/></svg>',
    list: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16v2H4V6Zm0 5h16v2H4v-2Zm0 5h16v2H4v-2Z" fill="currentColor"/></svg>',
    live: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 6a6 6 0 1 0 0 12 6 6 0 0 0 0-12Zm0 9a3 3 0 1 1 0-6 3 3 0 0 1 0 6Z" fill="currentColor"/></svg>',
    check: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 16.2 4.8 12l-1.4 1.4L9 19l12-12-1.4-1.4Z" fill="currentColor"/></svg>',
    report: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6Zm-1 7V3.5L18.5 9H13Z" fill="currentColor"/></svg>',
    hub: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4a3 3 0 0 1 3 3 3 3 0 0 1-1.5 2.6V11h2.5a3 3 0 0 1 3 3 3 3 0 0 1-2.6 3H14v2.5a3 3 0 1 1-4 0V17H7.6A3 3 0 0 1 5 14a3 3 0 0 1 3-3h2.5V9.6A3 3 0 0 1 9 7a3 3 0 0 1 3-3Z" fill="currentColor"/></svg>',
    layers: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3 1 9l11 6 11-6-11-6Zm0 9L3 7v2l9 5 9-5V7l-9 5Zm0 4L3 11v2l9 5 9-5v-2l-9 5Z" fill="currentColor"/></svg>',
    users: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 12a4 4 0 1 0 0-8 4 4 0 0 0 0 8Zm0 2c-3 0-9 1.5-9 4.5V21h18v-2.5C18 15.5 12 14 9 14Zm9.5-1c2 0 5.5 1 5.5 3v2h-5v-2c0-1.5-.5-2.4-1.5-3.4.3-.4.6-.4 1-.6Zm-.5-3a3.5 3.5 0 1 1 0-7 3.5 3.5 0 0 1 0 7Z" fill="currentColor"/></svg>',
    audit: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M19 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V5a2 2 0 0 0-2-2Zm-2 6h-4V7h4v2Zm0 4h-4v-2h4v2Zm0 4h-4v-2h4v2ZM7 7h4v10H7V7Z" fill="currentColor"/></svg>',
    settings: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M19.4 12.9a7 7 0 0 0 0-1.8l2-1.5-2-3.5-2.4 1a7 7 0 0 0-1.5-.9L15 4h-4l-.5 2.2c-.5.2-1 .5-1.5.9l-2.4-1-2 3.5 2 1.5a7 7 0 0 0 0 1.8l-2 1.5 2 3.5 2.4-1c.5.4 1 .7 1.5.9L11 20h4l.5-2.2c.5-.2 1-.5 1.5-.9l2.4 1 2-3.5-2-1.5ZM13 15a3 3 0 1 1 0-6 3 3 0 0 1 0 6Z" fill="currentColor"/></svg>',
    train: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 2C7 2 4 2.6 4 7v9.5C4 18.4 5.6 20 7.5 20L6 21.5V22h12v-.5L16.5 20a3.5 3.5 0 0 0 3.5-3.5V7c0-4.4-3-5-8-5Zm-6 9V7h12v4H6Zm2.5 6a1.5 1.5 0 1 1 0-3 1.5 1.5 0 0 1 0 3Zm7 0a1.5 1.5 0 1 1 0-3 1.5 1.5 0 0 1 0 3Z" fill="currentColor"/></svg>',
    silo: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 2h10v4H7V2Zm-1 5h12l1 4v9c0 1-.5 2-1.5 2h-11C5.5 22 5 21 5 20v-9l1-4Zm2 5v6h8v-6H8Z" fill="currentColor"/></svg>',
    ship: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 14l2 5.5a3 3 0 0 0 2.8 1.9h8.4A3 3 0 0 0 19 19.5L21 14h-2v-4a2 2 0 0 0-2-2h-1V6h-8v2H7a2 2 0 0 0-2 2v4H3Zm4-4h10v3.5l-5-1.4-5 1.4V10Z" fill="currentColor"/></svg>',
    ledger: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 3h12a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H7a4 4 0 0 1-4-4V5a2 2 0 0 1 2-2Zm0 14v0a2 2 0 0 0 2 2h10V5H5v12Zm3-9h8v2H8V8Zm0 4h8v2H8v-2Z" fill="currentColor"/></svg>'
  };

  const els = {
    view: document.getElementById("view"),
    title: document.getElementById("view-title"),
    section: document.getElementById("filter-section"),
    family: document.getElementById("filter-family"),
    quality: document.getElementById("filter-quality"),
    search: document.getElementById("filter-search"),
    globalSearch: document.getElementById("global-search"),
    globalDate: document.getElementById("global-date"),
    dataCount: document.getElementById("data-count"),
    filtersToolbar: document.getElementById("filters-toolbar"),
    periodBar: document.getElementById("period-bar"),
    periodToggle: document.getElementById("period-toggle"),
    periodSummary: document.getElementById("period-summary"),
    monthPicker: document.getElementById("month-picker"),
    rangePicker: document.getElementById("range-picker"),
    rangeFrom: document.getElementById("range-from"),
    rangeTo: document.getElementById("range-to"),
    calcModeToggle: document.getElementById("calc-mode-toggle"),
    liveIndicator: document.getElementById("live-indicator"),
    notificationCount: document.getElementById("notification-count"),
    notificationButton: document.getElementById("notification-button"),
    refreshButton: document.getElementById("refresh-button"),
    shiftName: document.getElementById("shift-name"),
    syncTime: document.getElementById("sync-time"),
    userAvatar: document.getElementById("user-avatar"),
    userName: document.getElementById("user-name"),
    userRole: document.getElementById("user-role"),
    userToggle: document.getElementById("user-toggle"),
    userMenu: document.getElementById("user-menu"),
    userChip: document.getElementById("user-chip")
  };

  function init() {
    if (!state.authenticated) {
      showLoginScreen();
      return;
    }
    bootApp();
  }

  function bootApp() {
    document.getElementById("login-shell")?.setAttribute("hidden", "");
    document.getElementById("app-shell")?.removeAttribute("hidden");
    state.view = currentUser().defaultView || "dashboard";
    paintNavIcons();
    populateFilters();
    bindShell();
    bindShellShortcuts();
    bindPeriodSelector();
    syncCalcModeControls();
    paintShift();
    paintProfile();
    paintPeriodSummary();
    setInterval(paintShift, 60000);
    startAutoRefresh();
    syncNavActiveState();
    render();
  }

  function showLoginScreen() {
    const shell = document.getElementById("login-shell");
    const app = document.getElementById("app-shell");
    if (!shell || !app) return;
    shell.removeAttribute("hidden");
    app.setAttribute("hidden", "");
    const periodEl = document.getElementById("login-period");
    if (periodEl) {
      const r = periodRange();
      periodEl.textContent = r.label;
    }
    const roles = document.getElementById("login-roles");
    if (!roles) return;
    roles.innerHTML = PROFILES.map((p) => `
      <button type="button" class="role-card" data-login-profile="${escapeAttr(p.id)}">
        <span class="role-avatar" style="background:${p.color}">${escapeHtml(p.initials)}</span>
        <div class="role-body">
          <strong>${escapeHtml(p.role)}</strong>
          <span>${escapeHtml(p.name)}</span>
          <p>${escapeHtml(p.scope)}</p>
        </div>
        <span class="role-arrow" aria-hidden="true">→</span>
      </button>
    `).join("");
    roles.querySelectorAll("[data-login-profile]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const profile = PROFILES.find((p) => p.id === btn.dataset.loginProfile);
        if (!profile) return;
        state.profile = profile;
        state.authenticated = true;
        saveProfile(profile.id);
        saveAuthFlag(true);
        bootApp();
      });
    });
  }

  function logout() {
    state.authenticated = false;
    saveAuthFlag(false);
    closeProfileMenu();
    showLoginScreen();
  }

  function bindPeriodSelector() {
    if (!els.periodToggle) return;

    populateMonthPicker();
    syncPeriodControls();

    els.periodToggle.querySelectorAll("[data-period]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const period = btn.dataset.period;
        state.period = period;
        if (period === "custom" && (!state.customFrom || !state.customTo)) {
          const today = analysisAnchorDate();
          const ago = new Date(today.getTime() - 30 * 86400000);
          state.customFrom = ago.toISOString().slice(0, 10);
          state.customTo = today.toISOString().slice(0, 10);
        }
        resetPagination();
        syncPeriodControls();
        paintPeriodSummary();
        render();
      });
    });

    els.monthPicker?.addEventListener("change", () => {
      state.selectedMonth = els.monthPicker.value;
      try { localStorage.setItem("trace-port-selected-month", state.selectedMonth); } catch {}
      state.analysisDate = monthEndDate(state.selectedMonth);
      try { localStorage.setItem("trace-port-analysis-date", state.analysisDate); } catch {}
      if (state.period !== "month") {
        state.period = "month";
        syncPeriodControls();
      }
      resetPagination();
      paintPeriodSummary();
      render();
    });

    els.globalDate?.addEventListener("change", () => {
      if (!els.globalDate.value) return;
      state.analysisDate = els.globalDate.value;
      const dateMonth = state.analysisDate.slice(0, 7);
      if (dateMonth) {
        state.selectedMonth = dateMonth;
        try { localStorage.setItem("trace-port-selected-month", state.selectedMonth); } catch {}
      }
      try { localStorage.setItem("trace-port-analysis-date", state.analysisDate); } catch {}
      resetPagination();
      populateMonthPicker();
      syncPeriodControls();
      paintPeriodSummary();
      render();
    });

    els.rangeFrom?.addEventListener("change", () => {
      state.customFrom = els.rangeFrom.value;
      if (state.period === "custom") {
        resetPagination();
        paintPeriodSummary();
        render();
      }
    });

    els.calcModeToggle?.querySelectorAll("[data-calc-mode]").forEach((btn) => {
      btn.addEventListener("click", () => {
        state.calcMode = btn.dataset.calcMode;
        try { localStorage.setItem("trace-port-calc-mode", state.calcMode); } catch {}
        syncCalcModeControls();
        addLog("Mode de calcul", state.calcMode, `Mode de calcul changé en « ${state.calcMode === "official" ? "Officiel Excel" : "Brut opérationnel"} ».`);
        render();
      });
    });
    els.rangeTo?.addEventListener("change", () => {
      state.customTo = els.rangeTo.value;
      if (state.period === "custom") {
        resetPagination();
        paintPeriodSummary();
        render();
      }
    });
  }

  function populateMonthPicker() {
    if (!els.monthPicker) return;
    // Build a list of months: dataset months + current month +/- 12 months
    const months = new Set();
    (DATA?.events || []).forEach((e) => {
      const k = String(e.start || "").slice(0, 7);
      if (k) months.add(k);
    });
    const now = new Date();
    for (let i = -6; i <= 6; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      months.add(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
    }
    const sorted = Array.from(months).sort();
    els.monthPicker.innerHTML = sorted.map((m) => {
      const [y, mm] = m.split("-").map(Number);
      const label = `${MONTH_NAMES[(mm || 1) - 1]} ${y}`;
      return `<option value="${escapeAttr(m)}" ${m === state.selectedMonth ? "selected" : ""}>${escapeHtml(label)}</option>`;
    }).join("");
  }

  function syncCalcModeControls() {
    if (!els.calcModeToggle) return;
    els.calcModeToggle.querySelectorAll("[data-calc-mode]").forEach((b) => {
      const active = b.dataset.calcMode === state.calcMode;
      b.classList.toggle("is-active", active);
      b.setAttribute("aria-selected", active);
    });
  }

  function syncPeriodControls() {
    if (!els.periodToggle) return;
    els.periodToggle.querySelectorAll("[data-period]").forEach((b) => {
      const active = b.dataset.period === state.period;
      b.classList.toggle("is-active", active);
      b.setAttribute("aria-selected", active);
    });
    if (els.monthPicker) {
      els.monthPicker.hidden = state.period !== "month";
      els.monthPicker.value = state.selectedMonth;
    }
    if (els.globalDate) {
      els.globalDate.value = state.analysisDate || defaultAnalysisDate();
    }
    if (els.rangePicker) {
      els.rangePicker.hidden = state.period !== "custom";
      if (state.customFrom && els.rangeFrom) els.rangeFrom.value = state.customFrom;
      if (state.customTo && els.rangeTo) els.rangeTo.value = state.customTo;
    }
  }

  function paintPeriodSummary() {
    if (!els.periodSummary) return;
    const r = periodRange();
    const events = getFilteredEvents();
    const hours = sum(events, "durationHours");
    els.periodSummary.innerHTML = `<strong>${escapeHtml(r.label)}</strong> · ${fmtNumber(events.length, 0)} arrêts · ${fmtHours(hours)} cumul`;

    const sourceMonthEl = document.getElementById("source-month");
    if (sourceMonthEl) sourceMonthEl.textContent = r.label;

    const dataCountEl = document.getElementById("data-count");
    if (dataCountEl) {
      const trains = getAllTrains().filter((t) => eventInPeriod({ start: t.day })).length;
      const ships = getAllShips().filter(isValidShip).filter((s) => eventInPeriod({ start: s.start })).length;
      const total = events.length;
      dataCountEl.textContent = total === 0 ? "Aucune donnée"
        : `${fmtNumber(total, 0)} arrêts · ${fmtNumber(trains, 0)} trains · ${fmtNumber(ships, 0)} navires`;
    }
  }

  function startAutoRefresh() {
    if (state.autoRefreshTimer) clearInterval(state.autoRefreshTimer);
    state.autoRefreshTimer = setInterval(() => {
      state.lastSync = new Date();
      paintShift();
      paintPeriodSummary();
      pulseLiveIndicator();
      if (state.view === "dashboard") {
        render();
      } else {
        updateNotificationBadge();
      }
    }, AUTO_REFRESH_MS);
  }

  function pulseLiveIndicator() {
    if (!els.liveIndicator) return;
    els.liveIndicator.classList.add("pulsing");
    setTimeout(() => els.liveIndicator?.classList.remove("pulsing"), 800);
  }

  function paintNavIcons() {
    document.querySelectorAll(".nav-item").forEach((button) => {
      const icon = button.dataset.icon;
      if (!icon || button.querySelector(".nav-icon")) return;
      const svg = NAV_ICONS[icon] || "";
      button.insertAdjacentHTML("afterbegin", `<span class="nav-icon" aria-hidden="true">${svg}</span>`);
    });
  }

  function bindShell() {
    document.querySelectorAll(".nav-item").forEach((button) => {
      button.addEventListener("click", () => {
        state.view = button.dataset.view;
        resetPagination();
        syncNavActiveState();
        render();
      });
    });

    [els.section, els.family, els.quality].forEach((select) => {
      select.addEventListener("change", () => {
        state.filters.section = els.section.value;
        state.filters.family = els.family.value;
        state.filters.quality = els.quality.value;
        resetPagination();
        render();
      });
    });

    let searchDebounce;
    const runSearch = (value) => {
      state.filters.search = value.trim().toLowerCase();
      resetPagination();
      render();
    };
    const queueSearch = (value) => {
      clearTimeout(searchDebounce);
      searchDebounce = setTimeout(() => runSearch(value), 220);
    };

    els.search.addEventListener("input", () => {
      if (els.globalSearch) els.globalSearch.value = els.search.value;
      queueSearch(els.search.value);
    });
    els.globalSearch?.addEventListener("input", () => {
      if (els.search) els.search.value = els.globalSearch.value;
      queueSearch(els.globalSearch.value);
    });

    document.getElementById("export-csv").addEventListener("click", exportEventsCsv);
    document.getElementById("export-json").addEventListener("click", exportSummaryJson);
  }

  function bindShellShortcuts() {
    els.refreshButton?.addEventListener("click", () => {
      state.lastSync = new Date();
      paintShift();
      render();
    });

    els.notificationButton?.addEventListener("click", () => {
      state.view = "validation";
      resetPagination();
      syncNavActiveState();
      render();
    });

    els.userToggle?.addEventListener("click", (event) => {
      event.stopPropagation();
      toggleProfileMenu();
    });

    document.addEventListener("click", (event) => {
      if (!els.userChip?.contains(event.target)) closeProfileMenu();
    });
  }

  function toggleProfileMenu(force) {
    const menu = els.userMenu;
    if (!menu) return;
    const open = force !== undefined ? force : menu.hasAttribute("hidden");
    if (open) {
      menu.innerHTML = `
        <span class="user-menu-section">Changer de profil</span>
        ${PROFILES.map((profile) => `
          <button type="button" role="menuitem" class="user-menu-item ${profile.id === state.profile.id ? "is-active" : ""}" data-profile="${escapeAttr(profile.id)}">
            <span class="user-menu-avatar" style="background:${profile.color}">${escapeHtml(profile.initials)}</span>
            <span class="user-menu-text">
              <strong>${escapeHtml(profile.name)}</strong>
              <span>${escapeHtml(profile.role)}</span>
              <em>${escapeHtml(profile.scope)}</em>
            </span>
          </button>
        `).join("")}
        <div class="user-menu-divider"></div>
        <button type="button" role="menuitem" class="user-menu-item user-menu-logout" data-logout>
          <span class="user-menu-avatar" style="background:#94a3b8">↩</span>
          <span class="user-menu-text">
            <strong>Se déconnecter</strong>
            <span>Retour à l'écran de sélection</span>
          </span>
        </button>
      `;
      menu.removeAttribute("hidden");
      els.userToggle?.setAttribute("aria-expanded", "true");
      menu.querySelectorAll("[data-profile]").forEach((btn) => {
        btn.addEventListener("click", () => switchProfile(btn.dataset.profile));
      });
      menu.querySelector("[data-logout]")?.addEventListener("click", logout);
    } else {
      menu.setAttribute("hidden", "");
      els.userToggle?.setAttribute("aria-expanded", "false");
    }
  }

  function closeProfileMenu() {
    toggleProfileMenu(false);
  }

  function switchProfile(id) {
    const profile = PROFILES.find((p) => p.id === id);
    if (!profile) return;
    state.profile = profile;
    saveProfile(id);
    state.view = profile.defaultView || "dashboard";
    resetPagination();
    paintProfile();
    syncNavActiveState();
    closeProfileMenu();
    addLog("Profil", profile.id, `Changement de profil actif → ${profile.role}.`);
    render();
  }

  function paintProfile() {
    const profile = currentUser();
    if (els.userName) els.userName.textContent = profile.name;
    if (els.userRole) els.userRole.textContent = profile.role;
    if (els.userAvatar) {
      els.userAvatar.textContent = profile.initials;
      els.userAvatar.style.background = profile.color;
    }
  }

  function paintShift() {
    if (!els.shiftName) return;
    const shift = currentShift();
    els.shiftName.textContent = shift.label;
    els.shiftName.dataset.tone = shift.tone;
    if (els.syncTime) {
      els.syncTime.textContent = relativeTimeFrom(state.lastSync);
    }
  }

  function currentShift() {
    const hour = new Date().getHours();
    if (hour >= 6 && hour < 14) return { label: "Quart matin · 06h-14h", tone: "morning", team: "Equipe A" };
    if (hour >= 14 && hour < 22) return { label: "Quart après-midi · 14h-22h", tone: "afternoon", team: "Equipe B" };
    return { label: "Quart nuit · 22h-06h", tone: "night", team: "Equipe C" };
  }

  function relativeTimeFrom(date) {
    if (!date) return "—";
    const diff = Date.now() - new Date(date).getTime();
    const minutes = Math.floor(diff / 60000);
    if (minutes < 1) return "à l'instant";
    if (minutes < 60) return `il y a ${minutes} min`;
    const hours = Math.floor(minutes / 60);
    if (hours < 24) return `il y a ${hours} h`;
    return new Date(date).toLocaleString("fr-FR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });
  }

  function updateNotificationBadge() {
    if (!els.notificationCount) return;
    const pending = decorateEvents(getAllEvents()).filter((event) => event.status === "pending").length;
    els.notificationCount.textContent = pending > 99 ? "99+" : String(pending);
    els.notificationCount.classList.toggle("is-hidden", pending === 0);
    els.notificationButton?.setAttribute("title", `${pending} arrêt${pending > 1 ? "s" : ""} en attente de validation`);
  }

  function updateFiltersVisibility() {
    if (!els.filtersToolbar) return;
    const showFilters = VIEWS_WITH_FILTERS.has(state.view);
    const showPeriod = VIEWS_WITH_PERIOD.has(state.view);
    els.filtersToolbar.classList.toggle("is-compact", !showFilters);
    els.filtersToolbar.classList.toggle("is-hidden", !showFilters && !showPeriod && state.view !== "reporting");
    document.querySelectorAll("#filters-toolbar > label").forEach((label) => {
      const isPeriod = label.classList.contains("top-control");
      label.style.display = isPeriod ? (showPeriod ? "" : "none") : (showFilters ? "" : "none");
    });
    const actions = els.filtersToolbar.querySelector(".actions");
    if (actions) actions.style.display = VIEWS_WITH_EXPORTS.has(state.view) ? "" : "none";
  }

  function populateFilters() {
    const events = getAllEvents();
    fillSelect(els.section, [
      ["all", "Tous"],
      ["__blank__", "Non affecté"],
      ...unique(events.map((e) => e.sectionKey).filter(Boolean)).map((value) => [value, value])
    ]);
    fillSelect(els.family, [
      ["all", "Toutes"],
      ...unique(events.map((e) => e.family).filter(Boolean)).map((value) => [value, value])
    ]);
    fillSelect(els.quality, [
      ["all", "Toutes"],
      ...unique(events.map((e) => e.quality).filter(Boolean)).map((value) => [value, value])
    ]);
  }

  function fillSelect(select, options) {
    select.innerHTML = options.map(([value, label]) => `<option value="${escapeAttr(value)}">${escapeHtml(label)}</option>`).join("");
  }

  function getLocalEvents() {
    try {
      return JSON.parse(localStorage.getItem(EVENTS_KEY) || "[]");
    } catch {
      return [];
    }
  }

  function saveLocalEvents(events) {
    localStorage.setItem(EVENTS_KEY, JSON.stringify(events));
  }

  function getAllEvents() {
    return [...DATA.events, ...getLocalEvents()];
  }

  function getLocalTrains() {
    try {
      return JSON.parse(localStorage.getItem(TRAINS_KEY) || "[]");
    } catch {
      return [];
    }
  }

  function saveLocalTrains(trains) {
    localStorage.setItem(TRAINS_KEY, JSON.stringify(trains));
  }

  function getAllTrains() {
    return [...DATA.trains, ...getLocalTrains()];
  }

  function getLocalShips() {
    try {
      return JSON.parse(localStorage.getItem(SHIPS_KEY) || "[]");
    } catch {
      return [];
    }
  }

  function saveLocalShips(ships) {
    localStorage.setItem(SHIPS_KEY, JSON.stringify(ships));
  }

  function getAllShips() {
    return [...DATA.ships, ...getLocalShips()];
  }

  function getValidationOverrides() {
    try {
      return JSON.parse(localStorage.getItem(VALIDATIONS_KEY) || "{}");
    } catch {
      return {};
    }
  }

  function saveValidationOverrides(overrides) {
    localStorage.setItem(VALIDATIONS_KEY, JSON.stringify(overrides));
  }

  function getLocalLogs() {
    try {
      return JSON.parse(localStorage.getItem(LOGS_KEY) || "[]");
    } catch {
      return [];
    }
  }

  function saveLocalLogs(logs) {
    localStorage.setItem(LOGS_KEY, JSON.stringify(logs));
  }

  function addLog(action, objectId, detail, user = currentUser().name) {
    const logs = getLocalLogs();
    logs.unshift({
      id: `LOG-${Date.now()}`,
      at: new Date().toISOString(),
      user,
      action,
      objectId,
      detail
    });
    saveLocalLogs(logs.slice(0, 300));
  }

  function getFilteredEvents() {
    const query = state.filters.search;
    return getAllEvents().filter((event) => {
      const sectionOk = state.filters.section === "all" ||
        (state.filters.section === "__blank__" ? !event.sectionKey : event.sectionKey === state.filters.section);
      const familyOk = state.filters.family === "all" || event.family === state.filters.family;
      const qualityOk = state.filters.quality === "all" || event.quality === state.filters.quality;
      const haystack = [
        event.id,
        event.sectionKey,
        event.subEquipment,
        event.family,
        event.description,
        event.assignment,
        event.quality,
        event.destination
      ].join(" ").toLowerCase();
      const periodOk = eventInPeriod(event);
      return sectionOk && familyOk && qualityOk && periodOk && (!query || haystack.includes(query));
    });
  }

  function getEventsInRange(from, to) {
    if (!from || !to) return getAllEvents();
    return getAllEvents().filter((event) => {
      const ts = new Date(event.start || event.end || 0).getTime();
      return Number.isFinite(ts) && ts >= from.getTime() && ts <= to.getTime();
    });
  }

  function getAnalysisEvents() {
    return getFilteredEvents().filter((event) => {
      const decorated = decorateEvent(event);
      if (String(event.id || "").startsWith("LOCAL-")) return decorated.status === "validated";
      return decorated.status !== "rejected";
    });
  }

  /* === Official synthesis perimeter (matches Excel Synthèses sheet) === */

  const OFFICIAL_FAMILY_LOOKUP = (() => {
    const map = new Map();
    OFFICIAL_FAMILIES.forEach((fam) => map.set(normalize(fam), fam));
    Object.entries(FAMILY_ALIASES).forEach(([raw, official]) => {
      map.set(normalize(raw), official);
    });
    return map;
  })();

  function mapFamily(rawLabel) {
    if (rawLabel === null || rawLabel === undefined) return null;
    const raw = String(rawLabel).trim();
    if (!raw) return null;
    return OFFICIAL_FAMILY_LOOKUP.get(normalize(raw)) || null;
  }

  function inOfficialScope(event) {
    if (!event) return false;
    const sec = String(event.sectionKey || "").trim();
    if (!CHARGING_SECTIONS.includes(sec)) return false;
    return mapFamily(event.family) !== null;
  }

  function decorateWithOfficial(event) {
    const officialFamily = mapFamily(event.family);
    const sec = String(event.sectionKey || "").trim();
    const isOfficial = CHARGING_SECTIONS.includes(sec) && officialFamily !== null;
    return {
      ...event,
      _officialFamily: officialFamily,
      _isOfficial: isOfficial,
      _exclusionReason: isOfficial ? null
        : !CHARGING_SECTIONS.includes(sec)
          ? (event.sectionKey ? "Section hors périmètre chargement" : "Section non affectée")
          : "Famille hors taxonomie officielle"
    };
  }

  function getOfficialEvents() {
    return getFilteredEvents()
      .filter(inOfficialScope)
      .map((event) => ({
        ...event,
        family: mapFamily(event.family) // remap to canonical label so groupings align with Excel
      }));
  }

  function getOfficialAnalysisEvents() {
    return getAnalysisEvents()
      .filter(inOfficialScope)
      .map((event) => ({
        ...event,
        family: mapFamily(event.family)
      }));
  }

  function getUnmappedEvents() {
    return getFilteredEvents().filter((e) => !inOfficialScope(e)).map(decorateWithOfficial);
  }

  function getCalcEvents() {
    return state.calcMode === "raw" ? getAnalysisEvents() : getOfficialAnalysisEvents();
  }

  function buildDataQuality() {
    const all = getFilteredEvents();
    const official = getOfficialEvents();
    const unmapped = getUnmappedEvents();
    const officialHours = sum(official, "durationHours");
    const unmappedHours = sum(unmapped, "durationHours");
    const totalHours = officialHours + unmappedHours;

    const byUnmappedFamily = new Map();
    unmapped.forEach((e) => {
      const fam = (String(e.family || "").trim()) || "(vide)";
      const entry = byUnmappedFamily.get(fam) || {
        family: fam,
        hours: 0,
        count: 0,
        sections: new Set(),
        reason: e._exclusionReason
      };
      entry.hours += Number(e.durationHours) || 0;
      entry.count += 1;
      entry.sections.add(e.sectionKey || "—");
      byUnmappedFamily.set(fam, entry);
    });

    return {
      totalEvents: all.length,
      officialEvents: official.length,
      unmappedEvents: unmapped.length,
      totalHours,
      officialHours,
      unmappedHours,
      alignmentRatio: totalHours > 0 ? officialHours / totalHours : 1,
      unmappedFamilies: Array.from(byUnmappedFamily.values())
        .map((row) => ({
          family: row.family,
          hours: row.hours,
          count: row.count,
          sections: Array.from(row.sections).slice(0, 5).join(" · "),
          reason: row.reason,
          suggestion: suggestMapping(row.family)
        }))
        .sort((a, b) => b.hours - a.hours)
    };
  }

  function suggestMapping(raw) {
    const norm = normalize(raw);
    if (!norm || raw === "(vide)") return "Renseigner la famille avant validation";
    const direct = OFFICIAL_FAMILY_LOOKUP.get(norm);
    if (direct) return `Reclasser « ${direct} »`;
    for (const fam of OFFICIAL_FAMILIES) {
      const fn = normalize(fam);
      if (fn.startsWith(norm) || norm.startsWith(fn) || fn.includes(norm) || norm.includes(fn)) {
        return `Suggestion : « ${fam} »`;
      }
    }
    return "Hors taxonomie — ajouter au référentiel";
  }

  function buildReconciliation() {
    const official = getOfficialEvents();
    const byOfficial = new Map();
    OFFICIAL_FAMILIES.forEach((fam) => byOfficial.set(fam, 0));
    official.forEach((e) => {
      byOfficial.set(e.family, (byOfficial.get(e.family) || 0) + (Number(e.durationHours) || 0));
    });
    return OFFICIAL_FAMILIES.map((family) => {
      const hours = byOfficial.get(family) || 0;
      return { family, hours };
    });
  }

  function render() {
    if (VIEW_ALIASES[state.view]) state.view = VIEW_ALIASES[state.view];
    const title = VIEW_TITLES[state.view] || "Centre de commandement";
    els.title.textContent = title;

    const renderers = {
      dashboard: renderCommandCenter,
      entry: renderEntry,
      myStops: renderMyStops,
      currentStops: renderCurrentStopsView,
      validation: renderValidation,
      stopDetail: renderStopDetail,
      pareto: renderParetoAnalysis,
      performance: renderPerformanceCircuits,
      reporting: renderReportingHub,
      monthlySynth: renderMonthlySynthese,
      trains: renderTrainsView,
      stocks: renderStocksView,
      ships: renderShipsView,
      references: renderReferencesHub,
      reconciliation: renderReconciliationView,
      equipments: renderEquipments,
      stopNatures: renderStopNatures,
      users: renderUsers,
      settings: renderSettings,
      logs: renderLogs,
      daily: renderDailySynthesis,
      events: renderEvents,
      tonnage: renderTonnage,
      formulas: renderFormulas,
      dmaic: renderDmaic
    };
    updateFiltersVisibility();
    (renderers[state.view] || renderCommandCenter)();
    injectProcessContext();
    updateNotificationBadge();
    paintShift();
    paintPeriodSummary();
  }

  function injectProcessContext() {
    const [stage, process, nextAction, nextView] = VIEW_META[state.view] || VIEW_META.dashboard;
    const events = decorateEvents(getAllEvents());
    const pending = events.filter((event) => event.status === "pending").length;
    const validated = events.filter((event) => event.status === "validated").length;
    const critical = events.filter((event) => event.status === "pending" || Number(event.durationHours) >= 2).length;
    const nextLabel = VIEW_TITLES[nextView] || "Continuer";

    els.view.insertAdjacentHTML("afterbegin", `
      <section class="process-context">
        <div class="process-main">
          <span class="stage-badge">${escapeHtml(stage)}</span>
          <div>
            <h2>${escapeHtml(process)}</h2>
            <p>${escapeHtml(nextAction)}</p>
          </div>
        </div>
        <div class="process-flow" role="tablist" aria-label="Flux opérationnel">
          ${WORKFLOW_STEPS.map((step) => workflowStep(step.key, step.views.includes(state.view), step.target)).join("")}
        </div>
        <div class="process-actions">
          <div class="ops-counters">
            <span><strong>${fmtNumber(pending, 0)}</strong> attente</span>
            <span><strong>${fmtNumber(validated, 0)}</strong> validés</span>
            <span><strong>${fmtNumber(critical, 0)}</strong> critiques</span>
          </div>
          <button class="primary-button" type="button" data-target-view="${escapeAttr(nextView)}">${escapeHtml(nextLabel)}</button>
        </div>
      </section>
    `);
    bindQuickActions();
  }

  function workflowStep(label, active, target) {
    const targetLabel = VIEW_TITLES[target] || label;
    return `<button type="button" class="workflow-step${active ? " active" : ""}" data-target-view="${escapeAttr(target)}" role="tab" aria-selected="${active}" title="Aller à ${escapeAttr(targetLabel)}">${escapeHtml(label)}</button>`;
  }

  function renderCommandCenter() {
    const sourceEvents = getFilteredEvents();
    if (sourceEvents.length === 0) {
      els.view.innerHTML = renderPeriodEmptyState({
        icon: "chart",
        view: "dashboard",
        ctaLabel: "Justifier un premier arrêt",
        ctaTarget: "entry"
      });
      bindQuickActions();
      return;
    }
    const events = getCalcEvents();
    const metrics = computeMetrics(events);
    const pareto = topGroups(groupSum(events, "family"), 6);
    const trend = getAllDays().map(computeDaySummary).slice(-14);
    const alerts = buildDashboardAlerts(metrics, pareto, sourceEvents);
    const decorated = decorateEvents(sourceEvents);
    const quality = buildDataQuality();
    const pendingValidation = decorated.filter((e) => e.status === "pending").length;
    const criticalStops = decorated.filter((e) => Number(e.durationHours) >= 2);
    const longStops = decorated.filter((e) => e.status === "pending" && Number(e.durationHours) >= 2);
    const circuitRows = buildCircuitPerformance(metrics);
    const shift = currentShift();
    const profile = currentUser();
    const costOfDowntime = metrics.totalStopHours * COST_PER_STOP_HOUR_EUR;
    const trsDelta = metrics.trsGlobal - TRS_TARGET;
    const trsMaintDelta = metrics.trsMaintenance - TRS_MAINT_TARGET;
    const slaState = metrics.trsGlobal >= TRS_TARGET ? "ok" : metrics.trsGlobal >= TRS_TARGET - 0.05 ? "warn" : "alert";
    const actionQueue = buildActionQueue(decorated, metrics);
    const liveFeed = buildLiveFeed(decorated);
    const projection = buildProjection(metrics);
    const periodRangeData = periodRange();

    els.view.innerHTML = `
      <section class="command-hero">
        <div class="hero-greeting">
          <span class="hero-eyebrow">Bonjour ${escapeHtml(profile.name.split(" ")[0])}</span>
          <h2>${escapeHtml(profile.role)}</h2>
          <p>${escapeHtml(shift.label)} · ${escapeHtml(shift.team)} en service · ${escapeHtml(periodRangeData.label)}</p>
          <div class="hero-meta">
            <span class="meta-pill"><span class="dot dot-${slaState}"></span>SLA ${fmtPct(TRS_TARGET)} ${slaState === "ok" ? "respecté" : "à surveiller"}</span>
            <span class="meta-pill"><strong>${fmtHours(metrics.totalStopHours)}</strong> d'arrêts cumulés</span>
            <span class="meta-pill"><strong>${fmtNumber(pendingValidation, 0)}</strong> à valider</span>
            <span class="meta-pill"><strong>${fmtNumber(criticalStops.length, 0)}</strong> incidents critiques</span>
            ${projection.show ? `<span class="meta-pill projection ${projection.tone}">📈 Projection fin de mois : <strong>${fmtPct(projection.trsProjected)}</strong> ${projection.deltaLabel}</span>` : ""}
          </div>
        </div>
        <div class="hero-kpis">
          <article class="hero-kpi">
            <span class="hero-kpi-label">TRS Global</span>
            <strong class="hero-kpi-value tone-${slaState}">${fmtPct(metrics.trsGlobal)}</strong>
            <span class="hero-kpi-trend ${trsDelta >= 0 ? "up" : "down"}">${trsDelta >= 0 ? "▲" : "▼"} ${fmtPct(Math.abs(trsDelta))} vs objectif</span>
          </article>
          <article class="hero-kpi">
            <span class="hero-kpi-label">Coût indisponibilité</span>
            <strong class="hero-kpi-value" title="${fmtNumber(costOfDowntime, 0)} €">${fmtCompactEur(costOfDowntime)}</strong>
            <span class="hero-kpi-trend">${fmtNumber(COST_PER_STOP_HOUR_EUR, 0)} €/h × ${fmtHours(metrics.totalStopHours)}</span>
          </article>
          <article class="hero-kpi">
            <span class="hero-kpi-label">TRS Maintenance</span>
            <strong class="hero-kpi-value tone-${trsMaintDelta >= 0 ? "ok" : "warn"}">${fmtPct(metrics.trsMaintenance)}</strong>
            <span class="hero-kpi-trend ${trsMaintDelta >= 0 ? "up" : "down"}">Objectif ${fmtPct(TRS_MAINT_TARGET)}</span>
          </article>
          <article class="hero-kpi">
            <span class="hero-kpi-label">Débit moyen</span>
            <strong class="hero-kpi-value">${fmtNumber(metrics.cadenceTph, 0)} t/h</strong>
            <span class="hero-kpi-trend">${fmtNumber(metrics.pesageTotal, 0)} t pesées</span>
          </article>
        </div>
      </section>

      <div class="dashboard-main">
        <section class="panel action-queue-panel">
          <div class="panel-head">
            <div>
              <h2>Ce qui demande votre attention</h2>
              <p class="status-line">File d'actions priorisée selon votre rôle et la criticité opérationnelle.</p>
            </div>
            <span class="badge red">${actionQueue.length}</span>
          </div>
          ${renderActionQueue(actionQueue)}
        </section>

        <section class="panel">
          <div class="panel-head">
            <h2>Diagnostic Pareto — Top causes</h2>
            <span class="badge">${fmtHours(metrics.totalStopHours)}</span>
          </div>
          <div class="donut-layout">
            <canvas id="nature-donut" class="mini-chart donut-chart"></canvas>
            ${renderDonutLegend(pareto, metrics.totalStopHours)}
          </div>
        </section>

        <section class="panel">
          <div class="panel-head">
            <h2>Tendance des arrêts</h2>
            <span class="badge blue">14 jours glissants</span>
          </div>
          <canvas id="daily-trend-chart" class="chart"></canvas>
        </section>
      </div>

      <div class="metric-grid dashboard-kpis">
        ${kpiCard("TRS Global", fmtPct(metrics.trsGlobal), `Objectif ${fmtPct(TRS_TARGET)}`, slaState === "ok" ? "green" : slaState === "warn" ? "amber" : "red")}
        ${kpiCard("TRS Exploitation", fmtPct(metrics.trsExploitation), `Objectif ${fmtPct(TRS_TARGET)}`, "blue")}
        ${kpiCard("TRS Maintenance", fmtPct(metrics.trsMaintenance), `Objectif ${fmtPct(TRS_MAINT_TARGET)}`, "purple")}
        ${kpiCard("TRG", fmtPct(metrics.trgGlobal), "Disponibilité réelle", "teal")}
        ${kpiCard("MTTR", fmtHours(computeMttr(events)), "Temps moyen réparation", "amber")}
        ${kpiCard("MTBF", fmtHours(computeMtbf(events)), "Temps entre arrêts", "green")}
      </div>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>État des circuits manutention</h2>
            <p class="status-line">Vue temps réel par circuit : navire en cours, dernier arrêt, performance.</p>
          </div>
          <span class="badge cyan">CA30 · CB30 · CC30 · CD30</span>
        </div>
        <div class="circuit-grid">
          ${CIRCUITS.map((circuit) => renderCircuitStatusCard(circuit, decorated, metrics)).join("")}
        </div>
      </section>

      <div class="two-col">
        <section class="panel live-feed-panel">
          <div class="panel-head">
            <div>
              <h2>Flux opérationnel temps réel</h2>
              <p class="status-line">Stream des saisies les plus récentes — actualisé toutes les 30 s.</p>
            </div>
            <span class="live-badge"><span class="live-dot"></span>Live</span>
          </div>
          ${renderLiveFeed(liveFeed)}
        </section>

        <section class="panel">
          <div class="panel-head">
            <div>
              <h2>Cadence opérationnelle ${escapeHtml(periodRangeData.label)}</h2>
              <p class="status-line">Pace de la période vs objectif et projection automatique.</p>
            </div>
            <span class="badge ${projection.tone}">${projection.show ? "Pace " + fmtPct(projection.elapsed) : "Mois clos"}</span>
          </div>
          ${renderPaceTracker(metrics, projection)}
        </section>
      </div>

      <div class="dashboard-bottom">
        <section class="panel">
          <div class="panel-head">
            <h2>Performance par circuit</h2>
            <span class="badge cyan">TRS</span>
          </div>
          <canvas id="circuit-chart" class="mini-chart"></canvas>
        </section>
        <section class="panel">
          <div class="panel-head">
            <h2>Heatmap criticité</h2>
            <span class="badge red">Intensité</span>
          </div>
          ${renderOperationalHeatmap(events)}
        </section>
        <section class="panel">
          <div class="panel-head">
            <h2>Alertes opérationnelles</h2>
            <span class="badge red">${alerts.length}</span>
          </div>
          ${renderAlerts(alerts)}
        </section>
      </div>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Incidents critiques en cours</h2>
            <p class="status-line">Arrêts en attente d'action prioritaire — cliquer sur la ligne pour ouvrir la fiche.</p>
          </div>
          <span class="badge warn">${longStops.length} à traiter</span>
        </div>
        ${renderCurrentStops(longStops.slice(0, 6))}
      </section>

      ${renderDataQualityPanel(quality)}
    `;

    bindQuickActions();
    bindStopActions();
    requestAnimationFrame(() => {
      drawDonut("nature-donut", pareto, metrics.totalStopHours);
      drawDailyTrend("daily-trend-chart", trend);
      drawCircuitBars("circuit-chart", circuitRows);
    });
  }

  function renderDataQualityPanel(q) {
    const aligned = Math.round(q.alignmentRatio * 100);
    const tone = q.alignmentRatio >= 0.95 ? "ok" : q.alignmentRatio >= 0.8 ? "warn" : "alert";
    return `
      <section class="panel data-quality-panel">
        <div class="panel-head">
          <div>
            <h2>Qualité de donnée &amp; alignement Excel</h2>
            <p class="status-line">Comparaison entre le périmètre TRACE-PORT et la taxonomie officielle Excel <em>Synthèses</em>. Mode actif : <strong>${state.calcMode === "official" ? "Officiel Excel" : "Brut opérationnel"}</strong>.</p>
          </div>
          <button class="ghost-button" type="button" data-target-view="reconciliation">Voir la réconciliation complète →</button>
        </div>
        <div class="quality-grid">
          <article class="quality-card tone-${tone}">
            <span>Taux d'alignement Excel</span>
            <strong>${aligned} %</strong>
            <em>${fmtHours(q.officialHours)} alignés sur ${fmtHours(q.totalHours)} cumulés</em>
          </article>
          <article class="quality-card">
            <span>Arrêts officiels</span>
            <strong>${fmtNumber(q.officialEvents, 0)}</strong>
            <em>S/E ∈ CA30/CB30/CC30/CD30 · famille mappée</em>
          </article>
          <article class="quality-card ${q.unmappedEvents > 0 ? "tone-warn" : "tone-ok"}">
            <span>Arrêts non alignés</span>
            <strong>${fmtNumber(q.unmappedEvents, 0)}</strong>
            <em>${fmtHours(q.unmappedHours)} hors périmètre officiel</em>
          </article>
          <article class="quality-card">
            <span>Familles hors taxonomie</span>
            <strong>${fmtNumber(q.unmappedFamilies.length, 0)}</strong>
            <em>${q.unmappedFamilies.length > 0 ? "Action de classification recommandée" : "Aucune anomalie détectée"}</em>
          </article>
        </div>
        ${q.unmappedFamilies.length > 0 ? `
          <details class="quality-details" open>
            <summary>Détail des familles à classifier (${q.unmappedFamilies.length})</summary>
            <div class="table-wrap">
              <table class="data-table">
                <thead>
                  <tr><th>Famille raw</th><th>Heures</th><th>Arrêts</th><th>Sections</th><th>Motif</th><th>Suggestion</th></tr>
                </thead>
                <tbody>
                  ${q.unmappedFamilies.slice(0, 12).map((row) => `
                    <tr>
                      <td><strong>${escapeHtml(row.family)}</strong></td>
                      <td class="num">${fmtHours(row.hours)}</td>
                      <td class="num">${fmtNumber(row.count, 0)}</td>
                      <td>${escapeHtml(row.sections)}</td>
                      <td><span class="badge warn">${escapeHtml(row.reason || "—")}</span></td>
                      <td><em>${escapeHtml(row.suggestion)}</em></td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </details>
        ` : ""}
      </section>
    `;
  }

  function buildActionQueue(decorated, metrics) {
    const queue = [];
    const role = currentUser().role;
    const perms = currentRolePermissions();
    const pending = decorated.filter((e) => e.status === "pending");
    const critical = decorated.filter((e) => Number(e.durationHours) >= 2 && e.status !== "rejected").slice(0, 5);
    const longest = decorated.slice().sort((a, b) => (b.durationHours || 0) - (a.durationHours || 0)).slice(0, 5);

    if (perms.canValidate && pending.length) {
      queue.push({
        priority: "high",
        title: `${pending.length} arrêt${pending.length > 1 ? "s" : ""} en attente de validation`,
        detail: `Top ${Math.min(pending.length, 5)} à traiter en priorité — fiabilise le calcul TRS.`,
        cta: "Ouvrir la file",
        target: "validation"
      });
    }
    if (critical.length) {
      queue.push({
        priority: critical.length > 3 ? "high" : "med",
        title: `${critical.length} incident${critical.length > 1 ? "s" : ""} critique${critical.length > 1 ? "s" : ""} > 2 h`,
        detail: `Cause dominante : ${critical[0]?.family || "—"} sur ${critical[0]?.subEquipment || critical[0]?.sectionKey || "équipement"}.`,
        cta: "Voir en direct",
        target: "currentStops"
      });
    }
    if (metrics.trsMaintenance < TRS_MAINT_TARGET) {
      queue.push({
        priority: "med",
        title: `TRS Maintenance sous objectif (${fmtPct(metrics.trsMaintenance)})`,
        detail: `Objectif ${fmtPct(TRS_MAINT_TARGET)} — analyser les arrêts maintenance dominants.`,
        cta: "Analyse Pareto",
        target: "pareto"
      });
    }
    if (metrics.trsGlobal < TRS_TARGET - 0.02) {
      queue.push({
        priority: "high",
        title: `SLA TRS Global non atteint (${fmtPct(metrics.trsGlobal)})`,
        detail: `Écart ${fmtPct(Math.abs(metrics.trsGlobal - TRS_TARGET))} vs cible — diffuser un rapport d'écart.`,
        cta: "Générer rapport",
        target: "reporting"
      });
    }
    if (perms.canCreate && role === "Agent de quart") {
      queue.push({
        priority: "low",
        title: "Démarrer ma déclaration d'arrêt",
        detail: "Saisir un nouvel incident pour alimenter la traçabilité.",
        cta: "Nouvel arrêt",
        target: "entry"
      });
    }
    if (!queue.length) {
      queue.push({
        priority: "low",
        title: "Aucune action critique en attente",
        detail: "Continuer la surveillance opérationnelle.",
        cta: "Voir les incidents",
        target: "currentStops"
      });
    }
    return queue;
  }

  function renderActionQueue(queue) {
    return `
      <div class="action-queue">
        ${queue.map((item) => `
          <article class="action-item priority-${escapeAttr(item.priority)}">
            <span class="action-marker"></span>
            <div class="action-body">
              <strong>${escapeHtml(item.title)}</strong>
              <p>${escapeHtml(item.detail)}</p>
            </div>
            <button class="action-cta" type="button" data-target-view="${escapeAttr(item.target)}">${escapeHtml(item.cta)} →</button>
          </article>
        `).join("")}
      </div>
    `;
  }

  function buildLiveFeed(decorated) {
    const all = decorated.slice().sort((a, b) => {
      const ta = new Date(a.createdAt || a.end || a.start || 0).getTime();
      const tb = new Date(b.createdAt || b.end || b.start || 0).getTime();
      return tb - ta;
    });
    return all.slice(0, 8);
  }

  function renderLiveFeed(items) {
    if (!items.length) return `<div class="empty-state">Aucune saisie sur la période sélectionnée.</div>`;
    return `
      <div class="live-feed">
        ${items.map((item) => {
          const ts = new Date(item.createdAt || item.end || item.start || 0);
          const isLocal = String(item.id || "").startsWith("LOCAL-");
          const tone = item.status === "validated" ? "ok" : item.status === "rejected" ? "alert" : "warn";
          return `
            <article class="live-feed-item">
              <span class="live-feed-time">${escapeHtml(relativeTimeFrom(ts))}</span>
              <div class="live-feed-body">
                <strong>${escapeHtml(item.family || "Arrêt")} sur ${escapeHtml(item.sectionKey || "—")}</strong>
                <p>${escapeHtml(truncate(item.description || "—", 90))}</p>
                <span class="live-feed-meta">
                  <span class="badge ${tone}">${statusLabel(item.status)}</span>
                  <span>${fmtHours(item.durationHours)}</span>
                  <span>${escapeHtml(item.declaredBy || "—")}</span>
                  ${isLocal ? `<span class="badge cyan">Local</span>` : ""}
                </span>
              </div>
              <button class="ghost-button table-action" type="button" data-detail-id="${escapeAttr(item.id)}">Ouvrir</button>
            </article>
          `;
        }).join("")}
      </div>
    `;
  }

  function buildProjection(metrics) {
    const elapsed = periodElapsedFraction();
    if (state.period !== "mtd" || elapsed >= 0.999 || elapsed <= 0) {
      return { show: false, elapsed, trsProjected: metrics.trsGlobal, deltaLabel: "", tone: "" };
    }
    // Project total stop hours linearly to full period
    const projectedStopHours = metrics.totalStopHours / Math.max(elapsed, 0.01);
    const fullPeriodAvailable = metrics.chargingAvailableHours / Math.max(elapsed, 0.01);
    const trsProjected = Math.max(0, (fullPeriodAvailable - projectedStopHours) / fullPeriodAvailable);
    const delta = trsProjected - TRS_TARGET;
    const tone = trsProjected >= TRS_TARGET ? "green" : trsProjected >= TRS_TARGET - 0.05 ? "warn" : "red";
    const deltaLabel = `${delta >= 0 ? "+" : ""}${fmtPct(delta)} vs objectif`;
    return { show: true, elapsed, trsProjected, deltaLabel, tone, projectedStopHours };
  }

  function renderPaceTracker(metrics, projection) {
    if (!projection.show) {
      return `<div class="pace-tracker"><p class="status-line">Période figée — KPI consolidés.</p></div>`;
    }
    const elapsedPct = Math.max(0, Math.min(100, projection.elapsed * 100));
    const trsNow = metrics.trsGlobal;
    const trsTarget = TRS_TARGET;
    const stopBudget = metrics.chargingAvailableHours * (1 - TRS_TARGET);
    const stopUsed = metrics.totalStopHours;
    const stopBudgetPct = Math.max(0, Math.min(120, (stopUsed / Math.max(stopBudget, 1)) * 100));
    const budgetTone = stopBudgetPct <= elapsedPct + 5 ? "ok" : stopBudgetPct <= elapsedPct + 15 ? "warn" : "alert";

    return `
      <div class="pace-tracker">
        <div class="pace-row">
          <div class="pace-label">
            <span>Temps de la période écoulé</span>
            <strong>${fmtPct(projection.elapsed)}</strong>
          </div>
          <div class="pace-bar"><div class="pace-fill pace-elapsed" style="width:${elapsedPct}%"></div></div>
        </div>
        <div class="pace-row">
          <div class="pace-label">
            <span>Budget arrêts consommé (cible ${fmtPct(1 - TRS_TARGET)})</span>
            <strong class="tone-${budgetTone}">${fmtPct(stopUsed / Math.max(stopBudget, 1))}</strong>
          </div>
          <div class="pace-bar"><div class="pace-fill pace-budget tone-${budgetTone}" style="width:${Math.min(100, stopBudgetPct)}%"></div></div>
        </div>
        <div class="pace-projection">
          <div>
            <span>TRS actuel</span>
            <strong>${fmtPct(trsNow)}</strong>
          </div>
          <div>
            <span>Projection fin de mois</span>
            <strong class="tone-${projection.tone === "green" ? "ok" : projection.tone === "warn" ? "warn" : "alert"}">${fmtPct(projection.trsProjected)}</strong>
          </div>
          <div>
            <span>Objectif</span>
            <strong>${fmtPct(trsTarget)}</strong>
          </div>
          <div>
            <span>Arrêts restant tolérés</span>
            <strong>${fmtHours(Math.max(0, stopBudget - stopUsed))}</strong>
          </div>
        </div>
      </div>
    `;
  }

  function renderCircuitStatusCard(circuit, decorated, metrics) {
    const events = decorated.filter((e) => e.sectionKey === circuit.key);
    const lastEvent = events.slice().sort((a, b) => new Date(b.end || b.start || 0) - new Date(a.end || a.start || 0))[0];
    const stopHours = sum(events, "durationHours");
    const dayCount = Math.max(getAllDays().length, 1);
    const available = dayCount * 24;
    const trs = ratio(available - stopHours, available);
    const activeShip = lastEvent?.assignment || "—";
    const isRunning = !lastEvent || (lastEvent.end && new Date(lastEvent.end) < new Date(Date.now() - 30 * 60 * 1000));
    const tone = trs >= TRS_TARGET ? "ok" : trs >= TRS_TARGET - 0.05 ? "warn" : "alert";

    return `
      <article class="circuit-card tone-${tone}" style="--circuit-color:${circuit.color}">
        <header>
          <div class="circuit-id">
            <span class="circuit-dot" style="background:${circuit.color}"></span>
            <div>
              <strong>${escapeHtml(circuit.key)}</strong>
              <span>${escapeHtml(circuit.role)}</span>
            </div>
          </div>
          <span class="circuit-status ${isRunning ? "running" : "stopped"}">${isRunning ? "En marche" : "Arrêt"}</span>
        </header>
        <div class="circuit-kpi">
          <div><span>TRS</span><strong class="tone-${tone}">${fmtPct(trs)}</strong></div>
          <div><span>Arrêts cumul</span><strong>${fmtHours(stopHours)}</strong></div>
          <div><span>Événements</span><strong>${fmtNumber(events.length, 0)}</strong></div>
        </div>
        <footer>
          <span class="circuit-ship">Navire : <strong>${escapeHtml(activeShip)}</strong></span>
          ${lastEvent ? `<span class="circuit-last">Dernier arrêt : <strong>${escapeHtml(lastEvent.family || "—")}</strong> · ${fmtHours(lastEvent.durationHours)}</span>` : `<span class="circuit-last">Aucun arrêt</span>`}
        </footer>
      </article>
    `;
  }

  function renderDailySynthesis() {
    const days = getAllDays();
    if (!state.dailyDate || !days.includes(state.dailyDate)) {
      state.dailyDate = days[0] || dateKey(new Date().toISOString());
    }
    const summary = computeDaySummary(state.dailyDate);
    const topFamilies = topGroups(groupSum(summary.events, "family"), 8);
    const sectionRows = buildSynthesisRows(summary.events, CHARGING_SECTIONS, CHARGING_SECTIONS.length * 24);

    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Synthèse journalière</h2>
            <p class="status-line">Cette fiche se recalcule à partir des arrêts saisis pendant la journée, des trains reçus et des navires chargés.</p>
          </div>
          <label>Journée
            <select id="daily-date">
              ${days.map((day) => `<option value="${escapeAttr(day)}" ${day === state.dailyDate ? "selected" : ""}>${fmtDateFromKey(day)}</option>`).join("")}
            </select>
          </label>
        </div>
      </section>

      <div class="metric-grid">
        ${metric("Arrêts du jour", fmtHours(summary.stopHours), `${summary.events.length} événements`)}
        ${metric("Trains reçus", fmtNumber(summary.trainCount, 0), `${fmtNumber(summary.wagonCount, 0)} wagons, ${fmtNumber(summary.trainTonnage, 0)} t`)}
        ${metric("Navires", fmtNumber(summary.shipCount, 0), `${fmtNumber(summary.shipBascule, 0)} t bascule`)}
        ${metric("TRS global estimé", fmtPct(summary.trsGlobal), "Base CA/CB/CC/CD sur 24h")}
      </div>

      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Arrêts par S/E</h2>
            <span class="badge">${fmtDateFromKey(state.dailyDate)}</span>
          </div>
          ${renderSynthesisMini(sectionRows, CHARGING_SECTIONS.length * 24)}
        </section>
        <section class="panel">
          <div class="panel-head">
            <h2>Pareto du jour</h2>
            <span class="badge warn">${fmtHours(summary.stopHours)}</span>
          </div>
          ${renderProgressList(topFamilies, summary.stopHours)}
        </section>
      </div>

      <section class="panel">
        <div class="panel-head">
          <h2>Flux journalier</h2>
          <span class="badge cyan">Chargement + déchargement</span>
        </div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Bloc</th><th>Volume / nombre</th><th>Durée</th><th>Indicateur</th></tr></thead>
            <tbody>
              <tr><td>Tonnage pesage</td><td>${fmtNumber(summary.pesageTotal, 0)} t</td><td>-</td><td>Draft : ${fmtNumber(summary.draftTotal, 0)} t</td></tr>
              <tr><td>Déchargement trains</td><td>${fmtNumber(summary.trainTonnage, 0)} t</td><td>${fmtHours(summary.trainDurationHours)}</td><td>${fmtNumber(summary.trainCadence, 0)} t/h</td></tr>
              <tr><td>Chargement navires</td><td>${fmtNumber(summary.shipBascule, 0)} t</td><td>${fmtHours(summary.shipDurationHours)}</td><td>Ecart : ${fmtPct(summary.shipGapRatio)}</td></tr>
              <tr><td>Arrêts maintenance</td><td>${fmtHours(summary.maintenanceHours)}</td><td>-</td><td>TRS maintenance : ${fmtPct(summary.trsMaintenance)}</td></tr>
              <tr><td>Arrêts exploitation</td><td>${fmtHours(summary.exploitationHours)}</td><td>-</td><td>TRS exploitation : ${fmtPct(summary.trsExploitation)}</td></tr>
            </tbody>
          </table>
        </div>
      </section>

      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Trains du jour</h2>
            <span class="badge">${summary.trains.length}</span>
          </div>
          ${renderTrainsTable(summary.trains)}
        </section>
        <section class="panel">
          <div class="panel-head">
            <h2>Navires du jour</h2>
            <span class="badge">${summary.ships.length}</span>
          </div>
          ${renderShipsTable(summary.ships)}
        </section>
      </div>

      <section class="panel">
        <div class="panel-head">
          <h2>Détail des arrêts du jour</h2>
          <span class="badge">${summary.events.length}</span>
        </div>
        ${renderEventsTable(summary.events)}
      </section>
    `;

    document.getElementById("daily-date").addEventListener("change", (event) => {
      state.dailyDate = event.target.value;
      render();
    });
  }

  function renderMonthlySynthesis() {
    const events = getCalcEvents();
    const metrics = computeMetrics(events);
    const days = getAllDays();
    const dailyRows = days.map(computeDaySummary);
    const chargingRows = buildSynthesisRows(events, CHARGING_SECTIONS, metrics.chargingAvailableHours);
    const dischargeAvailable = Math.max(days.length * DISCHARGE_SECTIONS.length * 24, sum(getAllTrains(), "affectationHours"));
    const dischargeRows = buildSynthesisRows(events, DISCHARGE_SECTIONS, dischargeAvailable || 1);
    const familyRows = topGroups(groupSum(events, "family"), 40);
    const trains = getAllTrains();
    const ships = getAllShips();

    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Synthèse mensuelle</h2>
            <p class="status-line">Cette fiche correspond à la logique de la première feuille Excel envoyée en fin de mois aux services : elle consolide les arrêts journaliers, les tonnages, les trains et les navires.</p>
          </div>
          <span class="badge">Janvier 2026</span>
        </div>
      </section>

      <div class="metric-grid">
        ${metric("Total arrêts", fmtHours(metrics.totalStopHours), `${events.length} lignes Bilan`)}
        ${metric("Tonnage chargé", fmtNumber(metrics.pesageTotal, 0), "Pesage mensuel")}
        ${metric("Trains", fmtNumber(sum(trains, "trains"), 0), `${fmtNumber(sum(trains, "totalTonnage"), 0)} t déchargées`)}
        ${metric("Navires", fmtNumber(ships.length, 0), `${fmtNumber(sum(ships, "connaissement"), 0)} t connaissement`)}
      </div>

      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Chargement des navires</h2>
            <span class="badge">CA30 / CB30 / CC30 / CD30</span>
          </div>
          ${renderSynthesisMini(chargingRows, metrics.chargingAvailableHours)}
        </section>
        <section class="panel">
          <div class="panel-head">
            <h2>Déchargement des trains</h2>
            <span class="badge cyan">DA10 / DB10</span>
          </div>
          ${renderSynthesisMini(dischargeRows, dischargeAvailable || 1)}
        </section>
      </div>

      <section class="panel">
        <div class="panel-head">
          <h2>Pareto mensuel des arrêts</h2>
          <span class="badge warn">${fmtHours(metrics.totalStopHours)}</span>
        </div>
        ${renderProgressList(familyRows.slice(0, 15), metrics.totalStopHours)}
      </section>

      <section class="panel">
        <div class="panel-head">
          <h2>Synthèse par journée</h2>
          <span class="badge">${days.length} jours</span>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr><th>Jour</th><th>Arrêts</th><th>Trains</th><th>Tonnage trains</th><th>Navires</th><th>Bascule navires</th><th>Pesage</th><th>TRS global</th></tr>
            </thead>
            <tbody>
              ${dailyRows.map((row) => `
                <tr>
                  <td>${fmtDateFromKey(row.day)}</td>
                  <td>${fmtHours(row.stopHours)}</td>
                  <td>${fmtNumber(row.trainCount, 0)}</td>
                  <td>${fmtNumber(row.trainTonnage, 0)} t</td>
                  <td>${fmtNumber(row.shipCount, 0)}</td>
                  <td>${fmtNumber(row.shipBascule, 0)} t</td>
                  <td>${fmtNumber(row.pesageTotal, 0)} t</td>
                  <td>${fmtPct(row.trsGlobal)}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </section>
    `;
  }

  function renderEvents() {
    const events = getFilteredEvents();
    const byFamily = topGroups(groupSum(events, "family"), 10);
    const bySection = topGroups(groupSum(events, "sectionKey", "Non affecté"), 10);
    const pageInfo = paginate(events, state.pagination.events, PAGE_SIZE);

    els.view.innerHTML = `
      <div class="metric-grid">
        ${metric("Evénements", fmtNumber(events.length, 0), "Après filtres")}
        ${metric("Durée totale", fmtHours(sum(events, "durationHours")), "Somme des arrêts")}
        ${metric("Famille dominante", escapeHtml(byFamily[0]?.label || "-"), fmtHours(byFamily[0]?.value || 0))}
        ${metric("S/E dominant", escapeHtml(bySection[0]?.label || "-"), fmtHours(bySection[0]?.value || 0))}
      </div>
      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Journal des arrêts</h2>
            <span class="badge">${fmtNumber(events.length, 0)} lignes</span>
          </div>
          ${renderEventsTable(pageInfo.items)}
          ${renderPagination("events", pageInfo)}
        </section>
        <section class="panel">
          <h2>Répartition</h2>
          <canvas id="events-family-chart" class="mini-chart"></canvas>
          <div class="split-list">${renderProgressList(bySection, sum(events, "durationHours"))}</div>
        </section>
      </div>
    `;

    bindPagination();
    requestAnimationFrame(() => {
      drawBars("events-family-chart", byFamily, { color: "#0f766e", suffix: "h" });
    });
  }

  function renderMyStops() {
    const events = decorateEvents(getFilteredEvents());
    if (events.length === 0) {
      els.view.innerHTML = renderPeriodEmptyState({ icon: "stop", ctaLabel: "Créer un arrêt", ctaTarget: "entry" });
      bindQuickActions();
      return;
    }
    const myEvents = events.filter((event) => event.declaredBy === currentUser().name || String(event.id).startsWith("LOCAL-"));
    const rows = myEvents.length ? myEvents : events;
    const sorted = rows.slice().sort((a, b) => new Date(b.start || 0) - new Date(a.start || 0));
    const pending = sorted.filter((event) => event.status === "pending").length;
    const validated = sorted.filter((event) => event.status === "validated").length;
    const rejected = sorted.filter((event) => event.status === "rejected").length;
    const pageInfo = paginate(sorted, state.pagination.myStops, PAGE_SIZE);

    els.view.innerHTML = `
      <div class="metric-grid">
        ${metric("Mes arrêts", fmtNumber(sorted.length, 0), "Arrêts saisis ou affectés")}
        ${metric("En attente", fmtNumber(pending, 0), "À valider")}
        ${metric("Validés", fmtNumber(validated, 0), "Exploitables en synthèse")}
        ${metric("Rejetés", fmtNumber(rejected, 0), "À corriger")}
      </div>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Mes arrêts</h2>
            <p class="status-line">Table dynamique avec recherche globale, filtres de la barre supérieure, statut et accès au détail.</p>
          </div>
          <button class="primary-button" type="button" data-target-view="entry">Nouvel arrêt</button>
        </div>
        ${renderOperationalStopsTable(pageInfo.items, { actions: true })}
        ${renderPagination("myStops", pageInfo)}
      </section>
    `;

    bindQuickActions();
    bindStopActions();
    bindPagination();
  }

  function renderCurrentStopsView() {
    const allEvents = decorateEvents(getFilteredEvents());
    if (allEvents.length === 0) {
      els.view.innerHTML = renderPeriodEmptyState({ icon: "stop", ctaLabel: "Créer un arrêt", ctaTarget: "entry" });
      bindQuickActions();
      return;
    }
    const events = allEvents
      .filter((event) => event.status === "pending" || Number(event.durationHours) >= 1)
      .sort((a, b) => (b.durationHours || 0) - (a.durationHours || 0));
    const critical = events.filter((event) => Number(event.durationHours) >= 2).length;
    const totalHours = sum(events, "durationHours");
    const pageInfo = paginate(events, state.pagination.currentStops, PAGE_SIZE);

    els.view.innerHTML = `
      <div class="metric-grid">
        ${metric("Arrêts en cours", fmtNumber(events.length, 0), "Arrêts ouverts ou critiques")}
        ${metric("Critiques", fmtNumber(critical, 0), "Durée supérieure à 2 h")}
        ${metric("Durée cumulée", fmtHours(totalHours), "Impact opérationnel")}
        ${metric("Action requise", fmtNumber(events.filter((e) => e.status === "pending").length, 0), "À surveiller")}
      </div>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Suivi temps réel des arrêts</h2>
            <p class="status-line">Vue opérationnelle pour le poste de commande : équipement, nature, durée, statut et action rapide.</p>
          </div>
          <span class="badge red">${fmtHours(totalHours)}</span>
        </div>
        ${renderOperationalStopsTable(pageInfo.items, { actions: true, validationActions: true })}
        ${renderPagination("currentStops", pageInfo)}
      </section>
    `;

    bindStopActions();
    bindPagination();
  }

  function renderValidation() {
    const events = decorateEvents(getFilteredEvents());
    if (events.length === 0) {
      els.view.innerHTML = renderPeriodEmptyState({ icon: "stop", ctaLabel: "Créer un arrêt", ctaTarget: "entry" });
      bindQuickActions();
      return;
    }
    const pending = events.filter((event) => event.status === "pending");
    const validated = events.filter((event) => event.status === "validated").length;
    const rejected = events.filter((event) => event.status === "rejected").length;
    const today = new Date().toISOString().slice(0, 10);
    const overrides = getValidationOverrides();
    const validatedToday = Object.values(overrides).filter((o) => o.status === "validated" && String(o.at || "").slice(0, 10) === today).length;
    const rejectedToday = Object.values(overrides).filter((o) => o.status === "rejected" && String(o.at || "").slice(0, 10) === today).length;
    const pageInfo = paginate(pending, state.pagination.validation, PAGE_SIZE);

    els.view.innerHTML = `
      <div class="metric-grid">
        ${metric("En attente", fmtNumber(pending.length, 0), "À traiter")}
        ${metric("Validés aujourd'hui", fmtNumber(validatedToday, 0), `Total cumulé : ${fmtNumber(validated, 0)}`)}
        ${metric("Rejetés aujourd'hui", fmtNumber(rejectedToday, 0), `Total cumulé : ${fmtNumber(rejected, 0)}`)}
        ${metric("Taux validation", fmtPct(ratio(validated, events.length || 1)), "Sur la base filtrée")}
      </div>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Validation des arrêts</h2>
            <p class="status-line">Valider, rejeter ou ouvrir le détail d'un arrêt. Chaque action alimente automatiquement les logs.</p>
          </div>
          <span class="badge warn">${pending.length} en attente</span>
        </div>
        ${renderOperationalStopsTable(pageInfo.items, { actions: true, validationActions: true })}
        ${renderPagination("validation", pageInfo)}
      </section>
    `;

    bindStopActions();
    bindPagination();
  }

  function renderStopDetail() {
    const event = decorateEvent(findEventById(state.selectedEventId) || getFilteredEvents()[0]);
    if (!event) {
      els.view.innerHTML = `<section class="panel"><div class="empty-state">Aucun arrêt sélectionné.</div></section>`;
      return;
    }
    const history = buildEventHistory(event);

    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Détail de l'arrêt - ${escapeHtml(event.id)}</h2>
            <p class="status-line">${escapeHtml(event.subEquipment || event.sectionKey || "Equipement non renseigné")} - ${escapeHtml(event.family || "Nature non renseignée")}</p>
          </div>
          <div class="inline-actions">
            <span class="status-pill ${statusTone(event.status)}">${statusLabel(event.status)}</span>
            <button class="ghost-button" type="button" data-print-detail>Imprimer</button>
          </div>
        </div>
      </section>

      <div class="detail-layout">
        <section class="panel">
          <h2>Informations arrêt</h2>
          <div class="detail-grid">
            ${detailItem("Equipement", event.subEquipment || event.sectionKey || "-")}
            ${detailItem("Circuit", circuitForSection(event.sectionKey))}
            ${detailItem("Nature d'arrêt", event.family || "-")}
            ${detailItem("Début", fmtDateTime(event.start))}
            ${detailItem("Fin", fmtDateTime(event.end))}
            ${detailItem("Durée", fmtHours(event.durationHours))}
            ${detailItem("Qualité", event.quality || "-")}
            ${detailItem("Affectation", event.assignment || "-")}
            ${detailItem("Déclaré par", event.declaredBy)}
            ${detailItem("Statut", statusLabel(event.status))}
          </div>
          <h3>Commentaire</h3>
          <p class="comment-box">${escapeHtml(event.description || "Aucun commentaire renseigné.")}</p>
          <div class="inline-actions">
            <button class="primary-button" type="button" data-validate-id="${escapeAttr(event.id)}">Valider</button>
            <button class="danger-button" type="button" data-reject-id="${escapeAttr(event.id)}">Rejeter</button>
          </div>
        </section>

        <section class="panel">
          <div class="panel-head">
            <h2>Historique</h2>
            <span class="badge">${history.length}</span>
          </div>
          <div class="timeline">
            ${history.map((item) => `
              <div class="timeline-item">
                <strong>${escapeHtml(item.action)}</strong>
                <span>${fmtDateTime(item.at)} - ${escapeHtml(item.user)}</span>
                <p>${escapeHtml(item.detail)}</p>
              </div>
            `).join("")}
          </div>
        </section>
      </div>
    `;

    bindStopActions();
    document.querySelector("[data-print-detail]")?.addEventListener("click", () => window.print());
  }

  function renderParetoAnalysis() {
    const events = getCalcEvents();
    if (events.length === 0) {
      els.view.innerHTML = renderPeriodEmptyState({ icon: "chart", ctaLabel: "Justifier un arrêt", ctaTarget: "entry" });
      bindQuickActions();
      return;
    }
    const pareto = topGroups(groupSum(events, "family"), 12);
    const total = sum(events, "durationHours");

    els.view.innerHTML = `
      <div class="metric-grid">
        ${metric("Causes analysées", fmtNumber(pareto.length, 0), "Top familles")}
        ${metric("Cause principale", escapeHtml(pareto[0]?.label || "-"), fmtHours(pareto[0]?.value || 0))}
        ${metric("Durée totale", fmtHours(total), "Base filtrée")}
        ${metric("Part Top 3", fmtPct(ratio(sum(pareto.slice(0, 3), "value"), total)), "Priorisation maintenance")}
      </div>
      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Répartition par nature</h2>
            <span class="badge">${fmtHours(total)}</span>
          </div>
          ${renderProgressList(pareto.slice(0, 8), total)}
        </section>
        <section class="panel">
          <div class="panel-head">
            <h2>Courbe de Pareto</h2>
            <span class="badge warn">% cumulé</span>
          </div>
          <canvas id="pareto-analysis-chart" class="chart"></canvas>
        </section>
      </div>
    `;

    requestAnimationFrame(() => drawPareto("pareto-analysis-chart", pareto, total));
  }

  function renderPerformanceCircuits() {
    const events = getCalcEvents();
    const metrics = computeMetrics(events);
    const circuits = buildCircuitPerformance(metrics);
    const chargingRows = buildSynthesisRows(events, CHARGING_SECTIONS, metrics.chargingAvailableHours);
    const dischargeRows = buildSynthesisRows(events, DISCHARGE_SECTIONS, Math.max(getAllDays().length * 48, 1));
    const equipmentRanking = buildEquipmentRanking(events, metrics);
    const periodHours = periodDays() * 24 * CHARGING_SECTIONS.length;

    els.view.innerHTML = `
      <div class="metric-grid">
        ${circuits.map((row, index) => kpiCard(row.label, fmtPct(row.value), "Performance circuit", ["green", "blue", "amber", "teal"][index] || "blue")).join("")}
      </div>

      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Performance circuits</h2>
            <span class="badge cyan">TRS · ${escapeHtml(periodRange().label)}</span>
          </div>
          <canvas id="performance-circuit-chart" class="chart"></canvas>
        </section>
        <section class="panel">
          <h2>Synthèse chargement</h2>
          ${renderSynthesisMini(chargingRows, metrics.chargingAvailableHours)}
        </section>
      </div>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Top équipements par impact</h2>
            <p class="status-line">Classement temps réel par durée d'arrêt cumulée sur la période — sous-équipements critiques en haut de liste.</p>
          </div>
          <span class="badge red">${fmtNumber(equipmentRanking.length, 0)} équipements actifs</span>
        </div>
        <div class="equipment-ranking">
          ${equipmentRanking.slice(0, 12).map((eq, idx) => `
            <div class="equipment-rank-row tone-${eq.tone}">
              <span class="rank-index">${idx + 1}</span>
              <div class="rank-body">
                <strong>${escapeHtml(eq.code)}</strong>
                <span class="status-line">${escapeHtml(eq.circuit)} · ${fmtNumber(eq.count, 0)} arrêts · cause dominante <em>${escapeHtml(eq.topFamily)}</em></span>
              </div>
              <div class="rank-bar"><div class="rank-fill" style="width:${eq.barPct}%"></div></div>
              <div class="rank-values">
                <span class="rank-time">${fmtHours(eq.hours)}</span>
                <span class="rank-avail">Dispo ${fmtPct(eq.availability)}</span>
              </div>
            </div>
          `).join("")}
        </div>
      </section>

      <section class="panel">
        <div class="panel-head">
          <h2>Synthèse déchargement trains</h2>
          <span class="badge">DA10 / DB10</span>
        </div>
        ${renderSynthesisMini(dischargeRows, Math.max(getAllDays().length * 48, 1))}
      </section>
    `;

    requestAnimationFrame(() => drawCircuitBars("performance-circuit-chart", circuits));
  }

  function buildEquipmentRanking(events, metrics) {
    const grouped = new Map();
    events.forEach((event) => {
      const code = event.subEquipment || event.sectionKey || "Non affecté";
      const current = grouped.get(code) || {
        code,
        circuit: circuitForSection(event.sectionKey || ""),
        hours: 0,
        count: 0,
        families: new Map()
      };
      current.hours += Number(event.durationHours) || 0;
      current.count += 1;
      const fam = event.family || "—";
      current.families.set(fam, (current.families.get(fam) || 0) + (Number(event.durationHours) || 0));
      grouped.set(code, current);
    });
    const rows = Array.from(grouped.values()).map((r) => {
      const topFamilyEntry = Array.from(r.families.entries()).sort((a, b) => b[1] - a[1])[0];
      const totalAvailable = Math.max(periodDays() * 24, 1);
      const availability = Math.max(0, Math.min(1, (totalAvailable - r.hours) / totalAvailable));
      const tone = availability >= 0.9 ? "ok" : availability >= 0.75 ? "warn" : "alert";
      return {
        code: r.code,
        circuit: r.circuit,
        hours: r.hours,
        count: r.count,
        topFamily: topFamilyEntry?.[0] || "—",
        availability,
        tone
      };
    }).sort((a, b) => b.hours - a.hours);
    const maxHours = rows[0]?.hours || 1;
    rows.forEach((r) => { r.barPct = Math.min(100, (r.hours / maxHours) * 100); });
    return rows;
  }

  function renderReportingHub() {
    const period = state.reportingPeriod === "monthly" ? "monthly" : "daily";
    const isMonthly = period === "monthly";
    const reportType = isMonthly ? "mensuel" : "journalier";
    const rows = buildReportRows(period);
    const events = getCalcEvents();
    const metrics = computeMetrics(events);

    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Centre de reporting &amp; diffusion</h2>
            <p class="status-line">Génération unifiée des rapports journaliers, mensuels et exports techniques — remplace l'envoi manuel des classeurs Excel.</p>
          </div>
          <div class="segmented-control" role="tablist" aria-label="Période de reporting">
            <button class="segment ${!isMonthly ? "is-active" : ""}" type="button" role="tab" aria-selected="${!isMonthly}" data-reporting-period="daily">Journalier</button>
            <button class="segment ${isMonthly ? "is-active" : ""}" type="button" role="tab" aria-selected="${isMonthly}" data-reporting-period="monthly">Mensuel</button>
          </div>
        </div>
        <div class="metric-grid">
          ${metric("TRS Global", fmtPct(metrics.trsGlobal), `Objectif ${fmtPct(TRS_TARGET)}`)}
          ${metric("Temps d'arrêt", fmtHours(metrics.totalStopHours), `${fmtNumber(events.length, 0)} événements`)}
          ${metric("Tonnage", `${fmtNumber(metrics.pesageTotal, 0)} t`, `Débit ${fmtNumber(metrics.cadenceTph, 0)} t/h`)}
          ${metric("Coût d'indisponibilité", `${fmtNumber(metrics.totalStopHours * COST_PER_STOP_HOUR_EUR, 0)} €`, "Estimation Janvier 2026")}
        </div>
      </section>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Rapports ${isMonthly ? "mensuels" : "journaliers"}</h2>
            <p class="status-line">Liasse standardisée : KPIs, Pareto, synthèses S/E, flux trains &amp; navires.</p>
          </div>
          <div class="inline-actions">
            <button class="primary-button" type="button" data-report-kind="${period}" data-report-format="html">Générer PDF</button>
            <button class="ghost-button" type="button" data-report-kind="${period}" data-report-format="csv">Générer Excel</button>
          </div>
        </div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Nom du rapport</th><th>Type</th><th>Période</th><th>Généré le</th><th>Actions</th></tr></thead>
            <tbody>
              ${rows.map((row) => `
                <tr>
                  <td>${escapeHtml(row.name)}</td>
                  <td>${escapeHtml(reportType)}</td>
                  <td>${escapeHtml(row.period)}</td>
                  <td>${fmtDateTime(row.generatedAt)}</td>
                  <td>
                    <button class="table-action" type="button" data-report-kind="${period}" data-report-format="html">PDF</button>
                    <button class="table-action" type="button" data-report-kind="${period}" data-report-format="csv">Excel</button>
                  </td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </section>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Exports techniques</h2>
            <p class="status-line">Pour intégration BI, audits qualité et archivage SMQE.</p>
          </div>
          <span class="badge">${fmtNumber(getAllEvents().length, 0)} arrêts disponibles</span>
        </div>
        <div class="export-grid">
          ${exportCard("Journal des arrêts", "CSV compatible Excel · 11 colonnes Bilan", "Exporter CSV", "events-csv")}
          ${exportCard("Synthèse opérationnelle", "JSON avec KPI, trains, navires, Pareto, synthèses journalières", "Exporter JSON", "summary-json")}
          ${exportCard("Rapport journalier (PDF)", "HTML imprimable, livré aux services chaque matin", "Générer", "report-daily")}
          ${exportCard("Rapport mensuel (PDF)", "Bilan officiel de fin de mois pour la direction", "Générer", "report-monthly")}
        </div>
      </section>
    `;

    bindReportButtons();
    bindExportCards();
    document.querySelectorAll("[data-reporting-period]").forEach((btn) => {
      btn.addEventListener("click", () => {
        state.reportingPeriod = btn.dataset.reportingPeriod;
        render();
      });
    });
  }

  function renderReconciliationView() {
    const recon = buildReconciliation();
    const quality = buildDataQuality();
    const useExcelRef = state.selectedMonth === "2026-01" || true; // Reference is January 2026
    const rows = recon.map((r) => {
      const expected = EXCEL_REFERENCE_JAN_2026[r.family];
      const diff = expected === undefined ? null : r.hours - expected;
      const absDiff = Math.abs(diff || 0);
      let status, statusTone;
      if (expected === undefined) {
        status = "Pas de référence Excel"; statusTone = "muted";
      } else if (absDiff < 0.1) {
        status = "OK"; statusTone = "ok";
      } else if (absDiff < 1.0) {
        status = "Écart mineur"; statusTone = "warn";
      } else {
        status = "Écart majeur"; statusTone = "alert";
      }
      return { ...r, expected, diff, status, statusTone };
    });
    const totalTP = sum(rows, "hours");
    const totalExcel = Object.values(EXCEL_REFERENCE_JAN_2026).reduce((a, b) => a + b, 0);
    const totalDiff = totalTP - totalExcel;
    const totalDiffTone = Math.abs(totalDiff) < 1 ? "ok" : Math.abs(totalDiff) < 10 ? "warn" : "alert";

    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Réconciliation Excel ↔ TRACE-PORT</h2>
            <p class="status-line">Audit de l'alignement entre la synthèse officielle <em>(Synthèses!row15)</em> et le calcul live TRACE-PORT, famille par famille. Référence : Janvier 2026.</p>
          </div>
          <button class="ghost-button" type="button" data-target-view="dashboard">← Retour au tableau de bord</button>
        </div>

        <div class="metric-grid">
          ${metric("Total Excel (officiel)", fmtHours(totalExcel), "Σ Synthèses!row15")}
          ${metric("Total TRACE-PORT (officiel)", fmtHours(totalTP), state.calcMode === "raw" ? "Mode brut sélectionné — basculer en officiel pour comparer" : "Mode officiel")}
          ${metric("Écart absolu", `${totalDiff >= 0 ? "+" : ""}${fmtHours(totalDiff)}`, `Tolérance : <0.1h = OK · <1h = mineur · ≥1h = majeur`)}
          ${metric("Alignement global", fmtPct(quality.alignmentRatio), `${fmtNumber(quality.officialEvents, 0)} arrêts officiels / ${fmtNumber(quality.totalEvents, 0)} total`)}
        </div>
      </section>

      <section class="panel">
        <div class="panel-head">
          <h2>Détail par famille</h2>
          <div class="reconciliation-legend">
            <span><span class="dot ok"></span> OK</span>
            <span><span class="dot warn"></span> Écart mineur</span>
            <span><span class="dot alert"></span> Écart majeur</span>
            <span><span class="dot muted"></span> Sans réf.</span>
          </div>
        </div>
        <div class="table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th>Famille officielle</th>
                <th>Excel (h)</th>
                <th>TRACE-PORT (h)</th>
                <th>Écart</th>
                <th>Statut</th>
              </tr>
            </thead>
            <tbody>
              ${rows.map((r) => `
                <tr class="recon-row tone-${r.statusTone}">
                  <td><strong>${escapeHtml(r.family)}</strong></td>
                  <td class="num">${r.expected !== undefined ? fmtHours(r.expected) : "—"}</td>
                  <td class="num">${fmtHours(r.hours)}</td>
                  <td class="num">${r.diff === null ? "—" : (r.diff >= 0 ? "+" : "") + fmtHours(r.diff)}</td>
                  <td><span class="status-pill ${r.statusTone === "ok" ? "green" : r.statusTone === "warn" ? "amber" : r.statusTone === "alert" ? "red" : ""}">${escapeHtml(r.status)}</span></td>
                </tr>
              `).join("")}
              <tr class="row-total">
                <td><strong>Total chargement</strong></td>
                <td class="num"><strong>${fmtHours(totalExcel)}</strong></td>
                <td class="num"><strong>${fmtHours(totalTP)}</strong></td>
                <td class="num"><strong>${totalDiff >= 0 ? "+" : ""}${fmtHours(totalDiff)}</strong></td>
                <td><span class="status-pill ${totalDiffTone === "ok" ? "green" : totalDiffTone === "warn" ? "amber" : "red"}">${totalDiffTone === "ok" ? "Aligné" : totalDiffTone === "warn" ? "Écart mineur" : "Écart majeur"}</span></td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      ${quality.unmappedFamilies.length > 0 ? `
        <section class="panel">
          <div class="panel-head">
            <h2>Catégories hors taxonomie officielle (mode brut)</h2>
            <span class="badge red">${quality.unmappedFamilies.length} familles</span>
          </div>
          <p class="status-line">Ces entrées <strong>ne sont pas incluses</strong> dans la synthèse officielle. Elles ne contribueront aux KPI que si elles sont reclassées vers une famille officielle.</p>
          <div class="table-wrap">
            <table class="data-table">
              <thead>
                <tr><th>Famille raw</th><th>Heures</th><th>Arrêts</th><th>Sections</th><th>Motif d'exclusion</th><th>Suggestion</th></tr>
              </thead>
              <tbody>
                ${quality.unmappedFamilies.map((row) => `
                  <tr>
                    <td><strong>${escapeHtml(row.family)}</strong></td>
                    <td class="num">${fmtHours(row.hours)}</td>
                    <td class="num">${fmtNumber(row.count, 0)}</td>
                    <td>${escapeHtml(row.sections)}</td>
                    <td><span class="badge warn">${escapeHtml(row.reason || "—")}</span></td>
                    <td><em>${escapeHtml(row.suggestion)}</em></td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </section>
      ` : ""}

      <section class="panel">
        <div class="panel-head">
          <h2>Documentation des formules</h2>
          <span class="badge cyan">Source de vérité</span>
        </div>
        <div class="formula-doc">
          ${Object.entries(KPI_DEFINITIONS).map(([key, def]) => `
            <article class="formula-doc-card">
              <header>
                <strong>${escapeHtml(def.label)}</strong>
                <span class="badge">${escapeHtml(def.sheet)}</span>
              </header>
              <code>${escapeHtml(def.formula)}</code>
            </article>
          `).join("")}
        </div>
      </section>
    `;

    bindQuickActions();
  }

  function renderReferencesHub() {
    const tab = state.referencesTab === "natures" ? "natures" : "equipments";
    const equipments = buildEquipmentRows();
    const totals = groupSum(getFilteredEvents(), "family");
    const families = unique([...DATA.families.map((f) => f.name), ...Array.from(totals.keys())]);

    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Référentiels industriels</h2>
            <p class="status-line">Maintenance de la cartographie : équipements, circuits et familles d'arrêt utilisées par les KPI.</p>
          </div>
          <div class="segmented-control">
            <button class="segment ${tab === "equipments" ? "is-active" : ""}" type="button" data-ref-tab="equipments">Équipements &amp; circuits (${equipments.length})</button>
            <button class="segment ${tab === "natures" ? "is-active" : ""}" type="button" data-ref-tab="natures">Natures d'arrêt (${families.length})</button>
          </div>
        </div>
      </section>
      ${tab === "equipments" ? renderEquipmentsPanel(equipments) : renderStopNaturesPanel(families, totals)}
    `;
    document.querySelectorAll("[data-ref-tab]").forEach((btn) => {
      btn.addEventListener("click", () => {
        state.referencesTab = btn.dataset.refTab;
        render();
      });
    });
    if (tab === "equipments") bindReferenceActions();
    else bindReferenceActions();
    bindPagination();
  }

  function renderEquipmentsPanel(equipments) {
    const pageInfo = paginate(equipments, state.pagination.equipments, PAGE_SIZE);
    return `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Équipements et circuits</h2>
            <p class="status-line">Référentiel extrait des S/E, sous-équipements et circuits.</p>
          </div>
          <button class="primary-button" type="button" data-reference-add="equipment">Ajouter</button>
        </div>
        ${equipments.length === 0 ? renderEmptyState({
          icon: "hub",
          title: "Aucun équipement référencé",
          message: "Importez le référentiel ou créez le premier équipement pour démarrer la cartographie industrielle.",
          ctaLabel: "Ajouter un équipement",
          ctaAction: "data-reference-add=\"equipment\""
        }) : `
          <div class="table-wrap">
            <table>
              <thead><tr><th>Code</th><th>Équipement</th><th>Circuit</th><th>Arrêts</th><th>Durée</th><th>Statut</th></tr></thead>
              <tbody>
                ${pageInfo.items.map((row) => `
                  <tr>
                    <td><strong>${escapeHtml(row.code)}</strong></td>
                    <td>${escapeHtml(row.name)}</td>
                    <td>${escapeHtml(row.circuit)}</td>
                    <td>${fmtNumber(row.count, 0)}</td>
                    <td>${fmtHours(row.hours)}</td>
                    <td><span class="status-pill green">Actif</span></td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
          ${renderPagination("equipments", pageInfo)}
        `}
      </section>
    `;
  }

  function renderStopNaturesPanel(families, totals) {
    const rows = families.map((family, index) => ({
      family,
      code: `NA-${String(index + 1).padStart(2, "0")}`,
      category: stopCategory(family),
      hours: totals.get(family) || 0,
      examples: DATA.families.find((item) => normalize(item.name) === normalize(family))?.examples || "-"
    }));
    const pageInfo = paginate(rows, state.pagination.stopNatures, PAGE_SIZE);
    return `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Natures d'arrêt</h2>
            <p class="status-line">Référentiel des familles utilisées dans les SUMIFS, le Pareto et les validations.</p>
          </div>
          <button class="primary-button" type="button" data-reference-add="nature">Ajouter</button>
        </div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Code</th><th>Libellé</th><th>Catégorie</th><th>Durée période</th><th>Exemples</th><th>Statut</th></tr></thead>
            <tbody>
              ${pageInfo.items.map((row) => `
                <tr>
                  <td><strong>${escapeHtml(row.code)}</strong></td>
                  <td>${escapeHtml(row.family)}</td>
                  <td>${escapeHtml(row.category)}</td>
                  <td>${fmtHours(row.hours)}</td>
                  <td>${escapeHtml(row.examples)}</td>
                  <td><button class="switch is-on" type="button" data-reference-toggle>Actif</button></td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
        ${renderPagination("stopNatures", pageInfo)}
      </section>
    `;
  }

  function renderEquipments() {
    const equipments = buildEquipmentRows();
    const pageInfo = paginate(equipments, state.pagination.equipments, PAGE_SIZE);
    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Équipements et circuits</h2>
            <p class="status-line">Référentiel opérationnel extrait des S/E, sous-équipements, trains et circuits de chargement/déchargement.</p>
          </div>
          <button class="primary-button" type="button" data-reference-add="equipment">Ajouter</button>
        </div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Code</th><th>Équipement</th><th>Circuit</th><th>Arrêts</th><th>Durée</th><th>Statut</th></tr></thead>
            <tbody>
              ${pageInfo.items.map((row) => `
                <tr>
                  <td>${escapeHtml(row.code)}</td>
                  <td>${escapeHtml(row.name)}</td>
                  <td>${escapeHtml(row.circuit)}</td>
                  <td>${fmtNumber(row.count, 0)}</td>
                  <td>${fmtHours(row.hours)}</td>
                  <td><span class="status-pill green">Actif</span></td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
        ${renderPagination("equipments", pageInfo)}
      </section>
    `;
    bindReferenceActions();
    bindPagination();
  }

  function renderStopNatures() {
    const totals = groupSum(getFilteredEvents(), "family");
    const families = unique([...DATA.families.map((f) => f.name), ...Array.from(totals.keys())]);
    const rows = families.map((family, index) => ({
      family,
      code: `NA-${String(index + 1).padStart(2, "0")}`,
      category: stopCategory(family),
      hours: totals.get(family) || 0,
      examples: DATA.families.find((item) => normalize(item.name) === normalize(family))?.examples || "-"
    }));
    const pageInfo = paginate(rows, state.pagination.stopNatures, PAGE_SIZE);
    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Natures d'arrêt</h2>
            <p class="status-line">Référentiel des familles utilisées dans les SUMIFS Excel, le Pareto et les validations.</p>
          </div>
          <button class="primary-button" type="button" data-reference-add="nature">Ajouter</button>
        </div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Code</th><th>Libellé</th><th>Catégorie</th><th>Durée mensuelle</th><th>Exemples</th><th>Actif</th></tr></thead>
            <tbody>
              ${pageInfo.items.map((row) => `
                <tr>
                  <td>${escapeHtml(row.code)}</td>
                  <td>${escapeHtml(row.family)}</td>
                  <td>${escapeHtml(row.category)}</td>
                  <td>${fmtHours(row.hours)}</td>
                  <td>${escapeHtml(row.examples)}</td>
                  <td><button class="switch is-on" type="button" data-reference-toggle>Actif</button></td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
        ${renderPagination("stopNatures", pageInfo)}
      </section>
    `;
    bindReferenceActions();
    bindPagination();
  }

  function renderUsers() {
    els.view.innerHTML = `
      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <div>
              <h2>Utilisateurs</h2>
              <p class="status-line">Gestion RBAC pour agent, chef d'équipe, superviseur, responsable et administrateur.</p>
            </div>
            <button class="primary-button" type="button" data-reference-add="user">Ajouter</button>
          </div>
          <div class="table-wrap">
            <table>
              <thead><tr><th>Nom</th><th>Prénom</th><th>Rôle</th><th>Service</th><th>Statut</th></tr></thead>
              <tbody>
                ${USERS.map((user) => `
                  <tr>
                    <td>${escapeHtml(user.lastName)}</td>
                    <td>${escapeHtml(user.firstName)}</td>
                    <td>${escapeHtml(user.role)}</td>
                    <td>${escapeHtml(user.service)}</td>
                    <td><span class="status-pill green">${escapeHtml(user.status)}</span></td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </section>
        <section class="panel">
          <h2>Rôles et permissions</h2>
          <div class="split-list">
            ${ROLES.map((role) => `
              <div class="list-row">
                <div>
                  <strong>${escapeHtml(role.name)}</strong>
                  <p class="status-line">${escapeHtml(role.permissions)}</p>
                </div>
                <span class="badge cyan">RBAC</span>
              </div>
            `).join("")}
          </div>
        </section>
      </div>
    `;
    bindReferenceActions();
  }

  function renderSettings() {
    const settings = [
      ["Période active", "Janvier 2026"],
      ["Objectif TRS global", "85 %"],
      ["Objectif TRS maintenance", "80 %"],
      ["Mode validation", "Chef d'équipe obligatoire"],
      ["Exports", "CSV, JSON, HTML imprimable"],
      ["Source initiale", DATA.sourceWorkbook]
    ];

    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Paramètres TRACE-PORT</h2>
            <p class="status-line">Paramétrage fonctionnel de la solution et rappel de l'architecture cible enterprise.</p>
          </div>
          <span class="badge cyan">RBAC + API REST + PostgreSQL</span>
        </div>
        <div class="requirement-grid">
          ${settings.map(([label, value]) => `
            <article class="requirement">
              <strong>${escapeHtml(label)}</strong>
              <p>${escapeHtml(value)}</p>
            </article>
          `).join("")}
        </div>
      </section>
      <section class="panel">
        <h2>Architecture 3 tiers cible</h2>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Couche</th><th>Technologie</th><th>Responsabilité</th></tr></thead>
            <tbody>
              <tr><td>Frontend</td><td>React / Next.js, TypeScript, TailwindCSS, Shadcn/UI, Recharts</td><td>Interfaces métier, dashboards, formulaires, filtres</td></tr>
              <tr><td>Backend</td><td>Node.js Express, API REST, JWT, RBAC</td><td>Contrôles métier, validation, KPI, rapports, sécurité</td></tr>
              <tr><td>Base de données</td><td>PostgreSQL</td><td>Arrêts, utilisateurs, référentiels, validations, logs et rapports</td></tr>
            </tbody>
          </table>
        </div>
      </section>
    `;
  }

  function renderLogs() {
    const logs = buildLogs();
    const pageInfo = paginate(logs, state.pagination.logs, PAGE_SIZE);
    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Historique des actions</h2>
            <p class="status-line">Traçabilité complète : connexions, créations, validations, rejets, exports et changements de statut.</p>
          </div>
          <span class="badge">${logs.length} logs</span>
        </div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Date / heure</th><th>Utilisateur</th><th>Action</th><th>Détail</th><th>Objet</th></tr></thead>
            <tbody>
              ${pageInfo.items.map((log) => `
                <tr>
                  <td>${fmtDateTime(log.at)}</td>
                  <td>${escapeHtml(log.user)}</td>
                  <td><span class="badge ${actionTone(log.action)}">${escapeHtml(log.action)}</span></td>
                  <td>${escapeHtml(log.detail)}</td>
                  <td>${escapeHtml(log.objectId || "-")}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
        ${renderPagination("logs", pageInfo)}
      </section>
    `;
    bindPagination();
  }

  function actionTone(action) {
    const value = String(action || "").toLowerCase();
    if (value.includes("validation") || value.includes("validé")) return "green";
    if (value.includes("rejet")) return "red";
    if (value.includes("création") || value.includes("creation")) return "blue";
    if (value.includes("export")) return "cyan";
    return "";
  }

  function renderEntry() {
    const familyOptions = unique([...DATA.families.map((f) => f.name), ...getAllEvents().map((e) => e.family).filter(Boolean)]);
    const sectionOptions = unique([...CHARGING_SECTIONS, ...DISCHARGE_SECTIONS, ...getAllEvents().map((e) => e.sectionKey).filter(Boolean)]);
    const qualityOptions = unique([...QUALITIES, ...Object.keys(DATA.tonnage[0]?.pesage || {}), ...getAllEvents().map((e) => e.quality).filter(Boolean)]);
    const ships = getAllShips();
    const activeShips = ships.slice().sort((a, b) => new Date(b.start || 0) - new Date(a.start || 0)).slice(0, 8);
    const shipOptions = unique([...activeShips.map((s) => s.name), ...getAllEvents().map((e) => e.assignment).filter(Boolean)]);
    const localCount = getLocalEvents().length;
    const profile = currentUser();

    els.view.innerHTML = `
      <section class="panel entry-context">
        <div class="entry-context-head">
          <span class="hero-eyebrow">Justification horaire d'arrêt</span>
          <h2>Capturer l'arrêt en cours sur les circuits concernés</h2>
          <p>Process Lean : remplacer la saisie manuelle dans le classeur Excel. Une seule saisie propage l'arrêt sur tous les circuits sélectionnés, conformément à la pratique opérationnelle (intempéries, marée, attente accostage = global, balance/LTE = par circuit).</p>
        </div>
        <div class="entry-context-meta">
          <span class="meta-pill"><span class="dot dot-${profile.id}"></span>${escapeHtml(profile.role)}</span>
          <span class="meta-pill"><strong>${localCount}</strong> saisies locales</span>
          <span class="meta-pill"><strong>${ships.length}</strong> navires actifs</span>
        </div>
      </section>

      <section class="panel">
        <div class="panel-head">
          <h2>Modèles d'arrêts courants</h2>
          <span class="badge cyan">Cliquer pour pré-remplir le formulaire</span>
        </div>
        <div class="template-grid">
          ${STOP_TEMPLATES.map((tpl, i) => `
            <button type="button" class="template-card scope-${escapeAttr(tpl.scope)}" data-template-index="${i}">
              <span class="template-family">${escapeHtml(tpl.family)}</span>
              <strong>${escapeHtml(tpl.desc)}</strong>
              <span class="template-meta">~${Math.floor(tpl.typical / 60)}h${tpl.typical % 60 ? " " + (tpl.typical % 60) + "min" : ""} · scope ${escapeHtml(scopeLabel(tpl.scope))}</span>
            </button>
          `).join("")}
        </div>
      </section>

      <section class="panel">
        <div class="panel-head">
          <h2>Saisie de l'arrêt</h2>
          <span class="badge warn">Formulaire SMQE compatible Bilan</span>
        </div>
        <form id="event-form" class="entry-form">
          <fieldset class="entry-circuits">
            <legend>Circuits affectés <em>(propagation parallèle)</em></legend>
            <div class="circuit-toggles">
              ${CIRCUITS.map((c) => `
                <label class="circuit-toggle" style="--circuit-color:${c.color}">
                  <input type="checkbox" name="circuits" value="${escapeAttr(c.key)}" checked>
                  <span class="toggle-mark"></span>
                  <span class="toggle-label">${escapeHtml(c.key)}</span>
                </label>
              `).join("")}
              ${DISCHARGE_SECTIONS.map((c) => `
                <label class="circuit-toggle is-discharge">
                  <input type="checkbox" name="circuits" value="${escapeAttr(c)}">
                  <span class="toggle-mark"></span>
                  <span class="toggle-label">${escapeHtml(c)}</span>
                </label>
              `).join("")}
            </div>
            <p class="entry-helper">Sélectionnez plusieurs circuits pour les causes globales (intempéries, marée, manque navire). Un arrêt sera créé pour chaque circuit.</p>
          </fieldset>

          <div class="entry-grid">
            <label>Sous-équipement<input name="subEquipment" placeholder="PD10, RC134, ..."></label>
            ${fieldSelect("family", "Famille d'arrêt", familyOptions)}
            <label>Début<input name="start" type="datetime-local" required></label>
            <label>Fin<input name="end" type="datetime-local" required></label>
            <label>Durée calculée<input id="duration-preview" readonly value="0 h"></label>
            <label>Navire affecté
              <input list="ship-options" name="assignment" placeholder="Sélectionner ou saisir">
              <datalist id="ship-options">${shipOptions.map((s) => `<option value="${escapeAttr(s)}">`).join("")}</datalist>
            </label>
            ${fieldSelect("quality", "Qualité", qualityOptions)}
            <label>Destination<input name="destination" placeholder="Destination produit"></label>
            <label class="full">Description / justification<textarea name="description" placeholder="Nature de l'anomalie, action terrain entreprise" rows="3"></textarea></label>
          </div>
          <aside class="impact-preview" id="impact-preview" aria-live="polite">
            <span class="impact-title">Impact estimé sur les KPI</span>
            <div class="impact-grid">
              <div><span>Circuits propagés</span><strong id="impact-circuits">${CHARGING_SECTIONS.length}</strong></div>
              <div><span>Heures d'arrêt ajoutées</span><strong id="impact-hours">0 h</strong></div>
              <div><span>Impact TRS (période)</span><strong id="impact-trs">—</strong></div>
              <div><span>Coût ajouté</span><strong id="impact-cost">0 €</strong></div>
            </div>
            <p class="impact-note" id="impact-note">Ajustez les circuits et la durée pour voir l'impact en direct.</p>
          </aside>

          <div class="entry-actions">
            <button class="primary-button" type="submit">Enregistrer &amp; propager</button>
            <button class="ghost-button" id="clear-template" type="button">Vider le formulaire</button>
            <button class="danger-button" id="clear-local" type="button">Réinitialiser saisies locales (${localCount})</button>
          </div>
        </form>
      </section>

      <section class="panel">
        <div class="panel-head">
          <h2>Dernières saisies locales</h2>
          <span class="badge">${localCount}</span>
        </div>
        ${renderEventsTable(getLocalEvents().slice().reverse().slice(0, 20))}
      </section>
    `;

    const form = document.getElementById("event-form");
    const preview = document.getElementById("duration-preview");
    const updatePreview = () => {
      const start = form.elements.start.value;
      const end = form.elements.end.value;
      preview.value = fmtHours(hoursBetweenLocal(start, end));
      updateImpactPreview(form);
    };
    form.elements.start.addEventListener("input", updatePreview);
    form.elements.end.addEventListener("input", updatePreview);
    form.querySelectorAll("input[name=circuits]").forEach((box) => box.addEventListener("change", () => updateImpactPreview(form)));
    form.addEventListener("submit", handleEventSubmit);
    updateImpactPreview(form);

    document.getElementById("clear-template")?.addEventListener("click", () => {
      form.reset();
      preview.value = "0 h";
      form.querySelectorAll("input[name=circuits]").forEach((box) => {
        box.checked = CHARGING_SECTIONS.includes(box.value);
      });
    });

    document.getElementById("clear-local")?.addEventListener("click", () => {
      if (!confirm("Effacer toutes les saisies locales ?")) return;
      saveLocalEvents([]);
      populateFilters();
      addLog("Saisie", "local-clear", "Réinitialisation des saisies locales d'arrêts.");
      render();
    });

    document.querySelectorAll("[data-template-index]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const tpl = STOP_TEMPLATES[Number(btn.dataset.templateIndex)];
        if (!tpl) return;
        form.elements.family.value = tpl.family;
        form.elements.description.value = tpl.desc;
        form.querySelectorAll("input[name=circuits]").forEach((box) => {
          box.checked = tpl.scope === "global" ? CHARGING_SECTIONS.includes(box.value)
            : tpl.scope === "stock" ? DISCHARGE_SECTIONS.includes(box.value)
            : tpl.scope === "circuit" ? box.value === "CA30"
            : CHARGING_SECTIONS.includes(box.value);
        });
        const now = new Date();
        const start = new Date(now.getTime() - tpl.typical * 60 * 1000);
        form.elements.start.value = formatLocalForInput(start);
        form.elements.end.value = formatLocalForInput(now);
        updatePreview();
        updateImpactPreview(form);
        form.elements.subEquipment.focus();
      });
    });
  }

  function scopeLabel(scope) {
    return {
      global: "Tous circuits",
      circuit: "Circuit isolé",
      ship: "Lié au navire",
      stock: "Silos / déchargement"
    }[scope] || scope;
  }

  function formatLocalForInput(date) {
    const pad = (n) => String(n).padStart(2, "0");
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
  }

  function updateImpactPreview(form) {
    if (!form) return;
    const circuitsEl = document.getElementById("impact-circuits");
    const hoursEl = document.getElementById("impact-hours");
    const trsEl = document.getElementById("impact-trs");
    const costEl = document.getElementById("impact-cost");
    const noteEl = document.getElementById("impact-note");
    if (!circuitsEl) return;

    const circuits = Array.from(form.querySelectorAll("input[name=circuits]:checked")).map((b) => b.value);
    const durationHours = hoursBetweenLocal(form.elements.start.value, form.elements.end.value);
    const addedStopHours = (durationHours || 0) * circuits.length;
    const metrics = computeMetrics(getAnalysisEvents());
    const currentTrs = metrics.trsGlobal;
    const available = Math.max(metrics.chargingAvailableHours, 1);
    const projectedStop = metrics.totalStopHours + addedStopHours;
    const projectedTrs = Math.max(0, (available - projectedStop) / available);
    const trsDelta = projectedTrs - currentTrs;
    const addedCost = addedStopHours * COST_PER_STOP_HOUR_EUR;

    circuitsEl.textContent = String(circuits.length);
    hoursEl.textContent = fmtHours(addedStopHours);
    trsEl.innerHTML = `<span class="tone-${trsDelta >= -0.005 ? "ok" : trsDelta >= -0.02 ? "warn" : "alert"}">${trsDelta >= 0 ? "+" : ""}${fmtPct(trsDelta)}</span>`;
    costEl.textContent = `${fmtNumber(addedCost, 0)} €`;

    if (!circuits.length) {
      noteEl.textContent = "Sélectionnez au moins un circuit affecté pour estimer l'impact.";
    } else if (!Number.isFinite(durationHours) || durationHours <= 0) {
      noteEl.textContent = `Renseignez début et fin — l'impact sera propagé sur ${circuits.length} circuit${circuits.length > 1 ? "s" : ""} (${circuits.join(" · ")}).`;
    } else {
      noteEl.textContent = `Cet arrêt ajoutera ${fmtHours(addedStopHours)} cumulés (${circuits.length} × ${fmtHours(durationHours)}). TRS Global passerait de ${fmtPct(currentTrs)} à ${fmtPct(projectedTrs)}.`;
    }
  }

  function renderTonnage() {
    const metrics = computeMetrics(getFilteredEvents());
    const qualityRows = Object.entries(metrics.qualityTotals).sort((a, b) => b[1] - a[1]);

    els.view.innerHTML = `
      <div class="metric-grid">
        ${metric("Pesage", fmtNumber(metrics.pesageTotal, 0), "Tonnage par qualités")}
        ${metric("Draft", fmtNumber(metrics.draftTotal, 0), "Tonnage draft")}
        ${metric("Bascule navires", fmtNumber(metrics.shipBascule, 0), "Navire!M")}
        ${metric("Ecart connaissement", fmtPct(metrics.shipGapRatio), "Navire!O")}
      </div>
      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Tonnage journalier</h2>
            <span class="badge">${DATA.tonnage.length} jours</span>
          </div>
          <canvas id="tonnage-chart" class="chart"></canvas>
        </section>
        <section class="panel">
          <h2>Qualités chargées</h2>
          ${renderProgressList(qualityRows.map(([label, value]) => ({ label, value })), metrics.pesageTotal)}
        </section>
      </div>
      <section class="panel">
        <div class="panel-head">
          <h2>Tableau tonnage</h2>
          <span class="badge">Pesage + draft</span>
        </div>
        ${renderTonnageTable()}
      </section>
    `;

    requestAnimationFrame(() => {
      drawLineBars("tonnage-chart", DATA.tonnage.map((d) => ({
        label: dayLabel(d.day),
        bar: d.pesageTotal || 0,
        line: d.draftTotal || 0
      })));
    });
  }

  /* ===== Operational chain: Trains, Stocks, Ships, Synthèse mensuelle ===== */

  function renderTrainsView() {
    const allTrains = getAllTrains();
    const trains = allTrains.filter((t) => eventInPeriod({ start: t.day }));
    const trainTotal = sum(trains, "totalTonnage");
    const wagons = sum(trains, "wagons");
    const trainCadence = average(trains.map((t) => t.cadenceTph).filter(Number.isFinite));
    const totalRetard = sum(trains, "delayHours");
    const localCount = getLocalTrains().length;
    const trsTrains = average(trains.map((t) => t.trsMaintenanceExploit).filter(Number.isFinite));
    const sortedTrains = trains.slice().sort((a, b) => new Date(b.day || 0) - new Date(a.day || 0));
    const recent = sortedTrains.slice(0, 31);

    els.view.innerHTML = `
      <section class="workflow-banner">
        <div class="workflow-step-banner is-active">
          <span class="banner-step-icon">1</span>
          <div><strong>Train</strong><span>Arrivée &amp; déchargement</span></div>
        </div>
        <div class="banner-arrow">→</div>
        <div class="workflow-step-banner">
          <span class="banner-step-icon">2</span>
          <div><strong>Silos DA / DB</strong><span>Stockage temporaire par qualité</span></div>
        </div>
        <div class="banner-arrow">→</div>
        <div class="workflow-step-banner">
          <span class="banner-step-icon">3</span>
          <div><strong>Circuits CA-CD</strong><span>Chargement navire</span></div>
        </div>
        <div class="banner-arrow">→</div>
        <div class="workflow-step-banner">
          <span class="banner-step-icon">4</span>
          <div><strong>Navire</strong><span>Pesage &amp; départ</span></div>
        </div>
      </section>

      <div class="metric-grid">
        ${metric("Trains reçus", fmtNumber(sum(trains, "trains"), 0), `${fmtNumber(wagons, 0)} wagons`)}
        ${metric("Tonnage déchargé", `${fmtNumber(trainTotal, 0)} t`, "Cumul mois")}
        ${metric("Cadence moyenne", `${fmtNumber(trainCadence, 0)} t/h`, "Vitesse déchargement")}
        ${metric("Retard cumulé", fmtHours(totalRetard), `TRS chaîne ${fmtPct(trsTrains || 0)}`)}
      </div>

      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <div>
              <h2>Enregistrer une arrivée train</h2>
              <p class="status-line">Capture le déchargement et alimente la traçabilité silo.</p>
            </div>
            <span class="badge cyan">${localCount} saisies locales</span>
          </div>
          <form id="train-form" class="form-grid">
            <label>Date arrivée<input name="day" type="date" required></label>
            <label>Nombre de trains<input name="trains" type="number" min="1" step="1" required></label>
            <label>Wagons<input name="wagons" type="number" min="0" step="1"></label>
            <label>Durée déchargement (h)<input name="durationHours" type="number" min="0" step="0.01"></label>
            <label>Tonnage silo DA (t)<input name="tonnageDA" type="number" min="0" step="0.01"></label>
            <label>Tonnage silo DB (t)<input name="tonnageDB" type="number" min="0" step="0.01"></label>
            <label>Bascule (t)<input name="tonnageBascule" type="number" min="0" step="0.01"></label>
            <label>Retard (h)<input name="delayHours" type="number" min="0" step="0.01"></label>
            <label class="wide">Observation<input name="observation" placeholder="Anomalie, semi-humide, silo cible"></label>
            <div class="inline-actions full">
              <button class="primary-button" type="submit">Enregistrer le train</button>
              <button class="danger-button" id="clear-local-trains" type="button">Réinitialiser saisies locales</button>
            </div>
          </form>
        </section>

        <section class="panel">
          <div class="panel-head">
            <h2>Cadence de déchargement</h2>
            <span class="badge cyan">t/h par jour</span>
          </div>
          <canvas id="train-cadence-chart" class="chart"></canvas>
        </section>
      </div>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Journal mensuel — déchargement trains</h2>
            <p class="status-line">Réplique de la feuille « Trains » du classeur Excel : tonnages silos, bascule, retard et cadence.</p>
          </div>
          <span class="badge">${recent.length} jour${recent.length > 1 ? "s" : ""} sur la période</span>
        </div>
        ${recent.length === 0 ? renderEmptyState({
          icon: "train",
          title: `Aucun train enregistré pour ${escapeHtml(periodRange().label)}`,
          message: "Aucune arrivée train n'a été enregistrée sur cette période. Sélectionnez un autre mois ou saisissez la première arrivée.",
          ctaLabel: "Enregistrer une arrivée train"
        }) : `
          <div class="table-wrap">
            <table class="data-table">
              <thead>
                <tr>
                  <th>Date</th>
                  <th>Trains</th>
                  <th>Wagons</th>
                  <th>Durée</th>
                  <th>Moy. / train</th>
                  <th>Silo DA</th>
                  <th>Silo DB</th>
                  <th>Bascule</th>
                  <th>Total</th>
                  <th>Affectation</th>
                  <th>Retard</th>
                  <th>Cadence</th>
                  <th>TRS</th>
                </tr>
              </thead>
              <tbody>
                ${recent.map((t) => `
                  <tr>
                    <td>${fmtDate(t.day)}</td>
                    <td class="num"><strong>${fmtNumber(t.trains, 0)}</strong></td>
                    <td class="num">${fmtNumber(t.wagons, 0)}</td>
                    <td class="num">${fmtHours(t.durationHours)}</td>
                    <td class="num">${fmtHours(t.averageHours)}</td>
                    <td class="num">${fmtNumber(t.tonnageDA || 0, 0)} t</td>
                    <td class="num">${fmtNumber(t.tonnageDB || 0, 0)} t</td>
                    <td class="num">${fmtNumber(t.tonnageBascule || 0, 0)} t</td>
                    <td class="num"><strong>${fmtNumber(t.totalTonnage || 0, 0)} t</strong></td>
                    <td class="num">${fmtHours(t.affectationHours)}</td>
                    <td class="num ${(t.delayHours || 0) > 0 ? "tone-red" : ""}">${fmtHours(t.delayHours)}</td>
                    <td class="num">${fmtNumber(t.cadenceTph, 0)} t/h</td>
                    <td class="num">${fmtPct(t.trsMaintenanceExploit || 0)}</td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        `}
      </section>
    `;

    bindFlowForms({ trainOnly: true });
    requestAnimationFrame(() => {
      drawBars("train-cadence-chart", trains.slice(-14).map((t) => ({
        label: dayLabel(t.day),
        value: t.cadenceTph || 0
      })), { color: "#1aa872", suffix: " t/h", yLabel: "Cadence (t/h)" });
    });
  }

  function renderStocksView() {
    const allTrains = getAllTrains();
    const allShips = getAllShips().filter(isValidShip);
    const trains = allTrains.filter((t) => eventInPeriod({ start: t.day }));
    const ships = allShips.filter((s) => eventInPeriod({ start: s.start }));
    const stockBalance = computeStockBalance();
    const totalIn = sum(trains, "totalTonnage");
    const totalOut = sum(ships, "bascule");
    const balance = totalIn - totalOut;
    const movements = buildStockMovements().filter((m) => eventInPeriod({ start: m.at })).slice(0, 12);

    els.view.innerHTML = `
      <section class="workflow-banner">
        <div class="workflow-step-banner">
          <span class="banner-step-icon">1</span>
          <div><strong>Train</strong><span>${fmtNumber(totalIn, 0)} t entrées</span></div>
        </div>
        <div class="banner-arrow">→</div>
        <div class="workflow-step-banner is-active">
          <span class="banner-step-icon">2</span>
          <div><strong>Silos DA / DB</strong><span>Solde ${fmtNumber(balance, 0)} t</span></div>
        </div>
        <div class="banner-arrow">→</div>
        <div class="workflow-step-banner">
          <span class="banner-step-icon">3</span>
          <div><strong>Circuits</strong><span>Vers chargement</span></div>
        </div>
        <div class="banner-arrow">→</div>
        <div class="workflow-step-banner">
          <span class="banner-step-icon">4</span>
          <div><strong>Navire</strong><span>${fmtNumber(totalOut, 0)} t sorties</span></div>
        </div>
      </section>

      <div class="metric-grid">
        ${metric("Entrées (trains)", `${fmtNumber(totalIn, 0)} t`, "Cumul mensuel via silos DA/DB")}
        ${metric("Sorties (navires)", `${fmtNumber(totalOut, 0)} t`, "Pesage bascule mensuel")}
        ${metric("Solde produit", `${fmtNumber(balance, 0)} t`, balance >= 0 ? "Stock disponible" : "Déficit théorique")}
        ${metric("Qualités actives", fmtNumber(QUALITIES.length, 0), QUALITIES.join(" · "))}
      </div>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Niveau des silos par qualité</h2>
            <p class="status-line">Estimation entrées train moins sorties pesage. Sert au pilotage de la disponibilité produit.</p>
          </div>
          <span class="badge cyan">Bilan mensuel</span>
        </div>
        <div class="silo-grid">
          ${SILOS.map((silo) => renderSiloCard(silo, stockBalance)).join("")}
        </div>
      </section>

      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Flux produit par qualité</h2>
            <span class="badge">Top entrées / sorties</span>
          </div>
          <div class="quality-flow">
            ${QUALITIES.map((q) => {
              const data = stockBalance.qualities[q] || { in: 0, out: 0 };
              const max = Math.max(...QUALITIES.map((qq) => Math.max(stockBalance.qualities[qq]?.in || 0, stockBalance.qualities[qq]?.out || 0)), 1);
              return `
                <div class="quality-row">
                  <span class="quality-tag">${escapeHtml(q)}</span>
                  <div class="quality-bars">
                    <div class="bar-track in"><div class="bar-fill" style="width:${(data.in / max) * 100}%"></div></div>
                    <div class="bar-track out"><div class="bar-fill" style="width:${(data.out / max) * 100}%"></div></div>
                  </div>
                  <div class="quality-values">
                    <span class="value-in">+${fmtNumber(data.in, 0)} t</span>
                    <span class="value-out">−${fmtNumber(data.out, 0)} t</span>
                  </div>
                </div>
              `;
            }).join("")}
          </div>
        </section>

        <section class="panel">
          <div class="panel-head">
            <h2>Mouvements récents</h2>
            <span class="badge cyan">Chronologique</span>
          </div>
          <div class="movement-list">
            ${movements.length ? movements.map((m) => `
              <div class="movement-item">
                <span class="movement-icon ${m.direction === "in" ? "in" : "out"}">${m.direction === "in" ? "↓" : "↑"}</span>
                <div>
                  <strong>${escapeHtml(m.label)}</strong>
                  <span class="status-line">${escapeHtml(m.detail)}</span>
                </div>
                <span class="movement-tonnage">${fmtNumber(m.tonnage, 0)} t</span>
              </div>
            `).join("") : `<div class="empty-state">Aucun mouvement enregistré.</div>`}
          </div>
        </section>
      </div>
    `;
  }

  function renderSiloCard(silo, balance) {
    const siloData = balance.silos[silo.key] || { in: 0, out: 0, qualities: {} };
    const net = siloData.in - siloData.out;
    const capacity = 50000;
    const fillPct = Math.max(0, Math.min(100, (net / capacity) * 100));
    const tone = fillPct >= 70 ? "high" : fillPct >= 30 ? "mid" : "low";
    const topQualities = Object.entries(siloData.qualities || {})
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3);

    return `
      <article class="silo-card tone-${tone}">
        <header>
          <div>
            <span class="silo-label">${escapeHtml(silo.label)}</span>
            <p>${escapeHtml(silo.role)}</p>
          </div>
          <span class="silo-tonnage">${fmtNumber(net, 0)} t</span>
        </header>
        <div class="silo-visual">
          <div class="silo-cylinder">
            <div class="silo-fill" style="height:${fillPct}%"></div>
            <span class="silo-percent">${fmtNumber(fillPct, 0)}%</span>
          </div>
        </div>
        <footer>
          <div class="silo-flow"><span class="flow-label in">Entrée</span><strong>+${fmtNumber(siloData.in, 0)} t</strong></div>
          <div class="silo-flow"><span class="flow-label out">Sortie</span><strong>−${fmtNumber(siloData.out, 0)} t</strong></div>
          ${topQualities.length ? `<div class="silo-qualities">Top qualités : ${topQualities.map(([q, v]) => `<span>${escapeHtml(q)} · ${fmtNumber(v, 0)}t</span>`).join(" ")}</div>` : ""}
        </footer>
      </article>
    `;
  }

  function computeStockBalance() {
    const balance = {
      silos: { DA: { in: 0, out: 0, qualities: {} }, DB: { in: 0, out: 0, qualities: {} } },
      qualities: {}
    };
    QUALITIES.forEach((q) => { balance.qualities[q] = { in: 0, out: 0 }; });

    getAllTrains().forEach((t) => {
      const da = Number(t.tonnageDA) || 0;
      const db = Number(t.tonnageDB) || 0;
      balance.silos.DA.in += da;
      balance.silos.DB.in += db;
    });

    DATA.tonnage.forEach((row) => {
      Object.entries(row.pesage || {}).forEach(([q, v]) => {
        if (!balance.qualities[q]) balance.qualities[q] = { in: 0, out: 0 };
        balance.qualities[q].in += v || 0;
      });
    });

    getAllShips().forEach((ship) => {
      const tot = Number(ship.bascule) || 0;
      balance.silos.DA.out += tot / 2;
      balance.silos.DB.out += tot / 2;
      if (ship.quality) {
        if (!balance.qualities[ship.quality]) balance.qualities[ship.quality] = { in: 0, out: 0 };
        balance.qualities[ship.quality].out += tot;
        const half = tot / 2;
        balance.silos.DA.qualities[ship.quality] = (balance.silos.DA.qualities[ship.quality] || 0) + half;
        balance.silos.DB.qualities[ship.quality] = (balance.silos.DB.qualities[ship.quality] || 0) + half;
      }
    });

    return balance;
  }

  function buildStockMovements() {
    const movements = [];
    getAllTrains().forEach((t) => {
      if (Number(t.totalTonnage) > 0) {
        movements.push({
          at: t.day,
          direction: "in",
          label: `Train ${t.trains || 1} rame${(t.trains || 1) > 1 ? "s" : ""} → Silos`,
          detail: `${fmtNumber(t.wagons || 0, 0)} wagons · DA ${fmtNumber(t.tonnageDA || 0, 0)} t / DB ${fmtNumber(t.tonnageDB || 0, 0)} t`,
          tonnage: t.totalTonnage
        });
      }
    });
    getAllShips().forEach((s) => {
      if (Number(s.bascule) > 0) {
        movements.push({
          at: s.start,
          direction: "out",
          label: `${s.name || "Navire"} (${s.quality || "—"})`,
          detail: `Poste ${s.berth || "—"} · connaissement ${fmtNumber(s.connaissement || 0, 0)} t · écart ${fmtPct(s.gapRatio || 0)}`,
          tonnage: s.bascule
        });
      }
    });
    return movements.sort((a, b) => new Date(b.at || 0) - new Date(a.at || 0));
  }

  function isValidShip(ship) {
    if (!ship) return false;
    const name = String(ship.name || "").trim();
    if (!name) return false;
    // Filter out structural rows from the Excel that aren't real ships
    if (/^(nombre navire|total|\d+)$/i.test(name)) return false;
    if (!ship.start) return false;
    return true;
  }

  function renderShipsView() {
    const allShips = getAllShips().filter(isValidShip);
    const ships = allShips.filter((s) => eventInPeriod({ start: s.start }));
    const sorted = ships.slice().sort((a, b) => new Date(b.start || 0) - new Date(a.start || 0));
    const totalBascule = sum(ships, "bascule");
    const totalConnaissement = sum(ships, "connaissement");
    const avgEcart = average(ships.map((s) => s.gapRatio).filter(Number.isFinite));
    const now = Date.now();
    const activeShips = sorted.filter((s) => {
      if (!s.start) return false;
      const start = new Date(s.start).getTime();
      const end = s.end ? new Date(s.end).getTime() : Infinity;
      return start <= now && now <= end;
    }).slice(0, 4);
    const qualityOptions = unique([...QUALITIES, ...ships.map((s) => s.quality).filter(Boolean)]);
    const localCount = getLocalShips().length;

    els.view.innerHTML = `
      <section class="workflow-banner">
        <div class="workflow-step-banner"><span class="banner-step-icon">1</span><div><strong>Train</strong><span>Phosphate déchargé</span></div></div>
        <div class="banner-arrow">→</div>
        <div class="workflow-step-banner"><span class="banner-step-icon">2</span><div><strong>Silos</strong><span>Stock par qualité</span></div></div>
        <div class="banner-arrow">→</div>
        <div class="workflow-step-banner"><span class="banner-step-icon">3</span><div><strong>Circuits CA-CD</strong><span>Reprise &amp; chargement</span></div></div>
        <div class="banner-arrow">→</div>
        <div class="workflow-step-banner is-active"><span class="banner-step-icon">4</span><div><strong>Navire</strong><span>Pesage &amp; départ</span></div></div>
      </section>

      <div class="metric-grid">
        ${metric("Navires chargés", fmtNumber(ships.length, 0), "Cumul mensuel")}
        ${metric("Tonnage bascule", `${fmtNumber(totalBascule, 0)} t`, "Pesage embarqué")}
        ${metric("Connaissement", `${fmtNumber(totalConnaissement, 0)} t`, "Conformité contractuelle")}
        ${metric("Écart moyen", fmtPct(avgEcart || 0), "Bascule vs connaissement")}
      </div>

      ${activeShips.length ? `
        <section class="panel">
          <div class="panel-head">
            <h2>Navires en cours de chargement</h2>
            <span class="badge cyan">${activeShips.length} actif${activeShips.length > 1 ? "s" : ""}</span>
          </div>
          <div class="ship-cards">
            ${activeShips.map(renderActiveShipCard).join("")}
          </div>
        </section>
      ` : `
        <section class="panel">
          <div class="panel-head">
            <h2>Navires en cours de chargement</h2>
            <span class="badge">Temps réel</span>
          </div>
          <div class="status-line empty-inline">Aucun navire en cours de chargement à l'instant. Les navires actifs apparaîtront ici lorsque les chargements seront en cours.</div>
        </section>
      `}

      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <div>
              <h2>Enregistrer un navire</h2>
              <p class="status-line">Saisie complète conforme à la feuille « Navire » du classeur Excel.</p>
            </div>
            <span class="badge cyan">${localCount} saisies locales</span>
          </div>
          <form id="ship-form" class="form-grid">
            <label>N° EC<input name="ecNumber" placeholder="Ex : 6617"></label>
            <label>Poste<input name="berth" placeholder="66"></label>
            <label>Navire<input name="name" placeholder="Nom du navire" required></label>
            ${fieldSelect("quality", "Qualité", qualityOptions)}
            <label>Début de chargement<input name="start" type="datetime-local" required></label>
            <label>Fin de chargement<input name="end" type="datetime-local" required></label>
            <label>Peseuse A (t)<input name="scaleA" type="number" min="0" step="0.01"></label>
            <label>Peseuse B (t)<input name="scaleB" type="number" min="0" step="0.01"></label>
            <label>Peseuse C (t)<input name="scaleC" type="number" min="0" step="0.01"></label>
            <label>Peseuse D (t)<input name="scaleD" type="number" min="0" step="0.01"></label>
            <label>Bascule totale (t)<input name="bascule" type="number" min="0" step="0.01"></label>
            <label>Connaissement (t)<input name="connaissement" type="number" min="0" step="0.01"></label>
            <label>Durée calculée<input id="ship-duration-preview" readonly value="0 h"></label>
            <label class="full">Observation<textarea name="observation" placeholder="Observation chargement"></textarea></label>
            <div class="inline-actions full">
              <button class="primary-button" type="submit">Enregistrer navire</button>
              <button class="danger-button" id="clear-local-ships" type="button">Réinitialiser saisies locales</button>
            </div>
          </form>
        </section>

        <section class="panel">
          <div class="panel-head">
            <h2>Écart connaissement vs bascule</h2>
            <span class="badge">% par navire</span>
          </div>
          <canvas id="ship-gap-chart" class="chart"></canvas>
        </section>
      </div>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Suivi de pesage statique — ${escapeHtml(periodRange().label)}</h2>
            <p class="status-line">Réplique de la feuille « Navire » : peseuses A à D, bascule, connaissement, écart.</p>
          </div>
          <span class="badge">${ships.length} navires</span>
        </div>
        ${sorted.length === 0 ? renderEmptyState({
          icon: "ship",
          title: "Aucun navire chargé sur cette période",
          message: "Aucun chargement n'a été enregistré pour la période sélectionnée. Sélectionnez un autre mois ou créez le premier navire."
        }) : `
          <div class="table-wrap">
            <table class="data-table">
              <thead>
                <tr>
                  <th>N°</th><th>Poste</th><th>Navire</th><th>Qualité</th><th>N° EC</th>
                  <th>Début</th><th>Fin</th><th>Durée</th>
                  <th>Pes. A</th><th>Pes. B</th><th>Pes. C</th><th>Pes. D</th>
                  <th>Bascule</th><th>Connaiss.</th><th>Écart</th>
                </tr>
              </thead>
              <tbody>
                ${sorted.map((s) => `
                  <tr>
                    <td class="num">${fmtNumber(s.number || 0, 0)}</td>
                    <td>${escapeHtml(s.berth || "—")}</td>
                    <td><strong>${escapeHtml(s.name || "—")}</strong></td>
                    <td><span class="badge cyan">${escapeHtml(s.quality || "—")}</span></td>
                    <td>${escapeHtml(s.ecNumber || "—")}</td>
                    <td>${fmtDateTime(s.start)}</td>
                    <td>${fmtDateTime(s.end)}</td>
                    <td class="num">${fmtHours(s.durationHours)}</td>
                    <td class="num">${fmtNumber(s.scaleA || 0, 0)}</td>
                    <td class="num">${fmtNumber(s.scaleB || 0, 0)}</td>
                    <td class="num">${fmtNumber(s.scaleC || 0, 0)}</td>
                    <td class="num">${fmtNumber(s.scaleD || 0, 0)}</td>
                    <td class="num"><strong>${fmtNumber(s.bascule || 0, 0)}</strong></td>
                    <td class="num">${fmtNumber(s.connaissement || 0, 0)}</td>
                    <td class="num ${(s.gapRatio || 0) > 0.005 ? "tone-amber" : "tone-green"}">${fmtPct(s.gapRatio || 0)}</td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        `}
      </section>
    `;

    bindFlowForms({ shipOnly: true });
    requestAnimationFrame(() => {
      drawBars("ship-gap-chart", ships.slice(-12).map((s) => ({
        label: truncate(s.name || "—", 10),
        value: Math.abs((s.gapRatio || 0) * 100)
      })), { color: "#7558e0", suffix: " %", yLabel: "Écart (%)" });
    });
  }

  function renderActiveShipCard(ship) {
    const stops = decorateEvents(getAllEvents()).filter((e) => e.assignment === ship.name);
    const lastStop = stops.sort((a, b) => new Date(b.end || b.start || 0) - new Date(a.end || a.start || 0))[0];
    const gapTone = (ship.gapRatio || 0) > 0.005 ? "warn" : "ok";
    const subtitle = [
      ship.berth ? `Poste ${ship.berth}` : null,
      ship.quality ? `qualité ${ship.quality}` : null,
      ship.ecNumber ? `EC ${ship.ecNumber}` : null
    ].filter(Boolean).join(" · ");
    return `
      <article class="ship-card">
        <header>
          <div>
            <strong>${escapeHtml(ship.name || "Navire")}</strong>
            <span class="status-line">${escapeHtml(subtitle || "Informations à compléter")}</span>
          </div>
          <span class="ship-status">En chargement</span>
        </header>
        <div class="ship-meta">
          <div><span>Bascule</span><strong>${fmtNumber(ship.bascule || 0, 0)} t</strong></div>
          <div><span>Connaiss.</span><strong>${fmtNumber(ship.connaissement || 0, 0)} t</strong></div>
          <div><span>Écart</span><strong class="tone-${gapTone}">${fmtPct(ship.gapRatio || 0)}</strong></div>
          <div><span>Durée</span><strong>${fmtHours(ship.durationHours)}</strong></div>
        </div>
        ${lastStop ? `<p class="ship-last-stop">Dernier arrêt : <strong>${escapeHtml(lastStop.family || "—")}</strong> sur ${escapeHtml(lastStop.sectionKey || "—")} · ${fmtHours(lastStop.durationHours)}</p>` : ""}
      </article>
    `;
  }

  function renderMonthlySynthese() {
    // Synthèse mensuelle reproduces the Excel "Synthèses" sheet — always use official perimeter
    const events = getOfficialAnalysisEvents();
    const metrics = computeMetrics(events);
    const families = OFFICIAL_FAMILIES;
    const chargingMatrix = CHARGING_SECTIONS.map((sectionKey) => {
      const row = { sectionKey, total: 0, percent: 0, families: {} };
      families.forEach((fam) => {
        const hours = events.filter((e) => e.sectionKey === sectionKey && e.family === fam).reduce((a, e) => a + (Number(e.durationHours) || 0), 0);
        row.families[fam] = hours;
        row.total += hours;
      });
      return row;
    });
    const grandTotal = chargingMatrix.reduce((a, r) => a + r.total, 0) || 1;
    chargingMatrix.forEach((r) => { r.percent = r.total / grandTotal; });
    const familyTotals = {};
    families.forEach((fam) => {
      familyTotals[fam] = chargingMatrix.reduce((a, r) => a + (r.families[fam] || 0), 0);
    });
    const ships = getAllShips();
    const trains = getAllTrains();

    els.view.innerHTML = `
      <section class="document-header">
        <div class="document-header-left">
          <span class="doc-ref">F-QE-721-01-02 · v X · Janvier 2026</span>
          <h2>Rapport mensuel d'activité — Manutention</h2>
          <p>Synthèse officielle livrée à la direction · OCP Port de Casablanca</p>
        </div>
        <div class="document-header-right">
          <button class="primary-button" type="button" data-report-kind="monthly" data-report-format="html">Imprimer / PDF</button>
          <button class="ghost-button" type="button" data-report-kind="monthly" data-report-format="csv">Exporter Excel</button>
        </div>
      </section>

      <section class="panel">
        <div class="panel-head">
          <h2>Chargement des navires — répartition des arrêts par S/E</h2>
          <span class="badge">${fmtHours(grandTotal)} total</span>
        </div>
        <div class="table-wrap synthesis-wrap">
          <table class="synthesis-table">
            <thead>
              <tr>
                <th class="sticky">S/E</th>
                <th>Total</th>
                <th>%</th>
                ${families.map((f) => `<th title="${escapeAttr(f)}">${escapeHtml(truncate(f, 14))}</th>`).join("")}
              </tr>
            </thead>
            <tbody>
              ${chargingMatrix.map((row) => `
                <tr>
                  <td class="sticky"><strong>${escapeHtml(row.sectionKey)}</strong></td>
                  <td><strong>${fmtHours(row.total)}</strong></td>
                  <td>${fmtPct(row.percent)}</td>
                  ${families.map((f) => {
                    const v = row.families[f] || 0;
                    return `<td class="${v > 0 ? "has-value" : ""}">${v > 0 ? fmtHours(v) : "—"}</td>`;
                  }).join("")}
                </tr>
              `).join("")}
              <tr class="row-total">
                <td class="sticky"><strong>Total chargement</strong></td>
                <td><strong>${fmtHours(grandTotal)}</strong></td>
                <td><strong>100%</strong></td>
                ${families.map((f) => {
                  const v = familyTotals[f] || 0;
                  return `<td>${v > 0 ? `<strong>${fmtHours(v)}</strong>` : "—"}</td>`;
                }).join("")}
              </tr>
              <tr class="row-percent">
                <td class="sticky">% total</td>
                <td>100%</td>
                <td>—</td>
                ${families.map((f) => {
                  const v = (familyTotals[f] || 0) / grandTotal;
                  return `<td>${v > 0 ? fmtPct(v) : "—"}</td>`;
                }).join("")}
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      <section class="panel">
        <div class="panel-head">
          <h2>Indicateurs mensuels consolidés</h2>
          <span class="badge cyan">KPI direction</span>
        </div>
        <div class="synthesis-kpis">
          ${synthKpi("Tonnage Chargé 24H", `${fmtNumber(metrics.pesageTotal, 0)} t`)}
          ${synthKpi("Tonnage Bascule", `${fmtNumber(metrics.shipBascule, 0)} t`)}
          ${synthKpi("Connaissement", `${fmtNumber(metrics.shipConnaissement, 0)} t`)}
          ${synthKpi("Nombre de navires", fmtNumber(ships.length, 0))}
          ${synthKpi("Débit / h", `${fmtNumber(metrics.cadenceTph, 0)} t/h`)}
          ${synthKpi("Cadence journalière 24h", `${fmtNumber(metrics.pesageTotal / Math.max(getAllDays().length, 1), 0)} t/j`)}
          ${synthKpi("Heure de marche", fmtHours(metrics.runningHours))}
          ${synthKpi("TRS exploitation", fmtPct(metrics.trsExploitation))}
          ${synthKpi("TRS maintenance", fmtPct(metrics.trsMaintenance))}
          ${synthKpi("TRS Global", fmtPct(metrics.trsGlobal))}
          ${synthKpi("TRG Global", fmtPct(metrics.trgGlobal))}
          ${synthKpi("Écart moyen", fmtPct(average(ships.map((s) => s.gapRatio).filter(Number.isFinite)) || 0))}
        </div>
      </section>

      <div class="two-col">
        <section class="panel">
          <div class="panel-head"><h2>Synthèse trains</h2><span class="badge cyan">${trains.length} jours</span></div>
          ${renderTrainsSummaryMini(trains)}
        </section>
        <section class="panel">
          <div class="panel-head"><h2>Synthèse navires</h2><span class="badge cyan">${ships.length} navires</span></div>
          ${renderShipsSummaryMini(ships)}
        </section>
      </div>
    `;

    bindReportButtons();
  }

  function synthKpi(label, value) {
    return `<article class="synth-kpi"><span>${escapeHtml(label)}</span><strong>${value}</strong></article>`;
  }

  function renderTrainsSummaryMini(trains) {
    const total = sum(trains, "totalTonnage");
    const wagons = sum(trains, "wagons");
    const delays = sum(trains, "delayHours");
    return `
      <div class="split-list">
        <div class="list-row"><strong>Trains reçus</strong><span>${fmtNumber(sum(trains, "trains"), 0)}</span></div>
        <div class="list-row"><strong>Wagons</strong><span>${fmtNumber(wagons, 0)}</span></div>
        <div class="list-row"><strong>Tonnage total</strong><span>${fmtNumber(total, 0)} t</span></div>
        <div class="list-row"><strong>Tonnage silo DA</strong><span>${fmtNumber(sum(trains, "tonnageDA"), 0)} t</span></div>
        <div class="list-row"><strong>Tonnage silo DB</strong><span>${fmtNumber(sum(trains, "tonnageDB"), 0)} t</span></div>
        <div class="list-row"><strong>Retard cumulé</strong><span>${fmtHours(delays)}</span></div>
      </div>
    `;
  }

  function renderShipsSummaryMini(ships) {
    const bascule = sum(ships, "bascule");
    const connaissement = sum(ships, "connaissement");
    const ecart = average(ships.map((s) => s.gapRatio).filter(Number.isFinite));
    return `
      <div class="split-list">
        <div class="list-row"><strong>Navires chargés</strong><span>${fmtNumber(ships.length, 0)}</span></div>
        <div class="list-row"><strong>Tonnage bascule</strong><span>${fmtNumber(bascule, 0)} t</span></div>
        <div class="list-row"><strong>Connaissement</strong><span>${fmtNumber(connaissement, 0)} t</span></div>
        <div class="list-row"><strong>Écart moyen</strong><span>${fmtPct(ecart || 0)}</span></div>
        <div class="list-row"><strong>Durée moyenne chargt</strong><span>${fmtHours(average(ships.map((s) => s.durationHours).filter(Number.isFinite)))}</span></div>
      </div>
    `;
  }

  function renderFlow() {
    const trains = getAllTrains();
    const ships = getAllShips();
    const qualityOptions = unique([...Object.keys(DATA.tonnage[0]?.pesage || {}), ...ships.map((s) => s.quality).filter(Boolean)]);
    const trainTotal = sum(trains, "totalTonnage");
    const wagons = sum(trains, "wagons");
    const shipsBascule = sum(ships, "bascule");
    const shipsConnaissement = sum(ships, "connaissement");
    const trainCadence = average(trains.map((t) => t.cadenceTph).filter(Number.isFinite));

    els.view.innerHTML = `
      <div class="metric-grid">
        ${metric("Trains", fmtNumber(sum(trains, "trains"), 0), `${fmtNumber(wagons, 0)} wagons`)}
        ${metric("Tonnage trains", fmtNumber(trainTotal, 0), "Total déchargé")}
        ${metric("Navires", fmtNumber(ships.length, 0), "Suivi de chargement")}
        ${metric("Connaissement", fmtNumber(shipsConnaissement, 0), `${fmtNumber(shipsBascule, 0)} bascule`)}
      </div>

      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Ajouter un train</h2>
            <span class="badge cyan">${getLocalTrains().length} saisies locales</span>
          </div>
          <form id="train-form" class="form-grid">
            <label>Date<input name="day" type="date" required></label>
            <label>Nombre de trains<input name="trains" type="number" min="1" step="1" required></label>
            <label>Wagons<input name="wagons" type="number" min="0" step="1"></label>
            <label>Tonnage total<input name="totalTonnage" type="number" min="0" step="0.01"></label>
            <label>Durée déchargement (h)<input name="durationHours" type="number" min="0" step="0.01"></label>
            <label>Retard (h)<input name="delayHours" type="number" min="0" step="0.01"></label>
            <label class="wide">Observation<input name="observation" placeholder="Retard, anomalie, silo, remarque"></label>
            <div class="inline-actions full">
              <button class="primary-button" type="submit">Ajouter train</button>
              <button class="danger-button" id="clear-local-trains" type="button">Réinitialiser trains locaux</button>
            </div>
          </form>
        </section>

        <section class="panel">
          <div class="panel-head">
            <h2>Ajouter un navire</h2>
            <span class="badge cyan">${getLocalShips().length} saisies locales</span>
          </div>
          <form id="ship-form" class="form-grid">
            <label>Navire<input name="name" placeholder="Nom du navire" required></label>
            ${fieldSelect("quality", "Qualité", qualityOptions)}
            <label>N° EC<input name="ecNumber" placeholder="EC"></label>
            <label>Début chargement<input name="start" type="datetime-local" required></label>
            <label>Fin chargement<input name="end" type="datetime-local" required></label>
            <label>Bascule<input name="bascule" type="number" min="0" step="0.01"></label>
            <label>Connaissement<input name="connaissement" type="number" min="0" step="0.01"></label>
            <label>Poste<input name="berth" placeholder="66"></label>
            <label>Durée calculée<input id="ship-duration-preview" readonly value="0 h"></label>
            <label class="full">Observation<textarea name="observation" placeholder="Observation chargement"></textarea></label>
            <div class="inline-actions full">
              <button class="primary-button" type="submit">Ajouter navire</button>
              <button class="danger-button" id="clear-local-ships" type="button">Réinitialiser navires locaux</button>
            </div>
          </form>
        </section>
      </div>

      <div class="two-col">
        <section class="panel">
          <h2>Cadence déchargement trains</h2>
          <canvas id="train-chart" class="chart"></canvas>
        </section>
        <section class="panel">
          <h2>Indicateurs trains</h2>
          ${renderProgressList([
            { label: "Cadence moyenne", value: trainCadence },
            { label: "Durée affectation", value: sum(trains, "affectationHours") },
            { label: "Retards", value: sum(trains, "delayHours") }
          ], Math.max(trainCadence, sum(trains, "affectationHours"), 1))}
        </section>
      </div>
      <section class="panel">
        <div class="panel-head">
          <h2>Navires chargés</h2>
          <span class="badge">${ships.length}</span>
        </div>
        ${renderShipsTable(ships)}
      </section>
      <section class="panel">
        <div class="panel-head">
          <h2>Déchargement trains</h2>
          <span class="badge">${trains.length} lignes saisies</span>
        </div>
        ${renderTrainsTable(trains)}
      </section>
    `;

    bindFlowForms();
    requestAnimationFrame(() => {
      drawBars("train-chart", trains.map((t) => ({ label: dayLabel(t.day), value: t.cadenceTph || 0 })), {
        color: "#1d7892",
        suffix: " t/h"
      });
    });
  }

  function renderFormulas() {
    const sheets = DATA.sheets.filter((s) => s.formulaCount > 0);
    const formulaRows = DATA.formulas.filter((f) => {
      const sheetOk = state.formulaSheet === "all" || f.sheet === state.formulaSheet;
      const q = state.formulaSearch.toLowerCase();
      const haystack = `${f.sheet} ${f.address} ${f.formula} ${f.cached}`.toLowerCase();
      return sheetOk && (!q || haystack.includes(q));
    });
    const requestRows = DATA.requests.slice(0, 120);

    els.view.innerHTML = `
      <div class="metric-grid">
        ${metric("Formules Excel", fmtNumber(DATA.formulasCount, 0), "Inventaire OOXML")}
        ${metric("Feuilles", fmtNumber(DATA.sheets.length, 0), "Classeur janvier")}
        ${metric("Requêtes intégrées", fmtNumber(DATA.requests.length, 0), "Feuilles EXPORTER / Feuil")}
        ${metric("Liens externes", fmtNumber(DATA.externalLinks.length, 0), "Sources historiques")}
      </div>
      <div class="formula-layout">
        <section class="panel">
          <h2>Moteur digital</h2>
          <div class="split-list">
            ${FORMULA_BLUEPRINTS.map((item) => `
              <div class="list-row">
                <div>
                  <strong>${escapeHtml(item.module)}</strong>
                  <p class="status-line">${escapeHtml(item.excel)}</p>
                  <p class="formula-code">${escapeHtml(item.digital)}</p>
                </div>
              </div>
            `).join("")}
          </div>
        </section>
        <section class="panel">
          <div class="panel-head">
            <h2>Audit des formules Excel</h2>
            <span class="badge">${fmtNumber(formulaRows.length, 0)}</span>
          </div>
          <div class="filters formula-search">
            <label>Feuille<select id="formula-sheet">
              <option value="all">Toutes</option>
              ${sheets.map((s) => `<option value="${escapeAttr(s.name)}" ${s.name === state.formulaSheet ? "selected" : ""}>${escapeHtml(s.name)}</option>`).join("")}
            </select></label>
            <label class="search-box">Recherche<input id="formula-search" type="search" value="${escapeAttr(state.formulaSearch)}" placeholder="SUMIFS, TRS, Bilan!"></label>
          </div>
          ${renderFormulaTable(formulaRows.slice(0, 300))}
        </section>
      </div>
      <section class="panel">
        <div class="panel-head">
          <h2>Feuilles et volumes</h2>
          <span class="badge">${DATA.sheets.length}</span>
        </div>
        ${renderSheetsTable()}
      </section>
      <section class="panel">
        <div class="panel-head">
          <h2>Requêtes / références importées</h2>
          <span class="badge">${DATA.requests.length}</span>
        </div>
        ${renderRequestsTable(requestRows)}
      </section>
      <section class="panel">
        <h2>Sources externes détectées</h2>
        <div class="pill-grid">
          ${DATA.externalLinks.map((link) => `<span class="pill">${escapeHtml(link)}</span>`).join("") || `<span class="pill">Aucune</span>`}
        </div>
      </section>
    `;

    document.getElementById("formula-search").addEventListener("input", (event) => {
      state.formulaSearch = event.target.value;
      renderFormulas();
    });
    document.getElementById("formula-sheet").addEventListener("change", (event) => {
      state.formulaSheet = event.target.value;
      renderFormulas();
    });
  }

  function renderDmaic() {
    const visionCards = [
      ["Problème actuel", "Fichiers Excel dispersés, traitements manuels, double saisie papier/Excel, erreurs de consolidation et indicateurs disponibles tardivement."],
      ["Vision cible", "Plateforme centralisée, automatisée et accessible aux acteurs du processus selon leurs rôles et leurs niveaux d'accès."],
      ["Objectif opérationnel", "Passer d'un suivi J+30 à une disponibilité quasi immédiate des arrêts, synthèses et KPI de manutention."],
      ["Valeur ajoutée", "Améliorer durablement la traçabilité, la fiabilité des données, la rapidité de traitement et la qualité de décision."]
    ];
    const roles = [
      ["Agent de quart", "Saisie directe des arrêts depuis un formulaire contrôlé."],
      ["Chef d'équipe", "Vérification, correction et validation des informations saisies."],
      ["Responsable exploitation", "Suivi temps réel des arrêts, circuits, tendances et alertes critiques."],
      ["Direction", "Pilotage global à travers les KPI, rapports et synthèses journalières ou mensuelles."]
    ];
    const leanGains = [
      "Suppression de la double saisie papier/Excel.",
      "Réduction des consolidations manuelles.",
      "Recalcul automatique des KPI après saisie ou validation.",
      "Historisation des créations, modifications, validations, rejets et suppressions.",
      "Disponibilité temps réel des informations pour accélérer la réaction terrain."
    ];
    const requirements = [
      ["Saisie des anomalies", "Formulaire structuré reprenant Bilan A:J, avec calcul automatique de la durée."],
      ["Synthèse mensuelle", "SUMIFS digitalisés par S/E et famille, totaux chargement et déchargement."],
      ["Performance manutention", "TRS exploitation, maintenance, global, TRG, cadence et heures de marche."],
      ["Traçabilité SMQE", "Historique horodaté, qualité, affectation, description, sources Excel et liens externes."],
      ["Analyse des causes", "Pareto dynamique par famille, S/E, qualité et description d'arrêt."],
      ["Flux logistique", "Croisement tonnage, trains, navires, bascule, connaissement et écarts."],
      ["Pilotage Poste de Commande", "Vue unique pour coordination installation, quai, maintenance et exploitation."],
      ["Amélioration DMAIC", "Base prête pour indicateurs Measure, priorisation Analyze, actions Improve."],
    ];

    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <h2>Présentation et vision</h2>
          <span class="badge">Système d'information centralisé</span>
        </div>
        <p class="status-line">
          TRACE-PORT modernise la traçabilité des arrêts de manutention au sein de la Direction Logistique Portuaire OCP Casablanca. La solution remplace un processus manuel, lent et fragmenté par une plateforme centralisée, contrôlée et orientée temps réel.
        </p>
      </section>

      <div class="requirement-grid">
        ${visionCards.map(([title, body]) => `
          <article class="requirement">
            <strong>${escapeHtml(title)}</strong>
            <p>${escapeHtml(body)}</p>
          </article>
        `).join("")}
      </div>

      <section class="panel">
        <div class="panel-head">
          <h2>Architecture fonctionnelle</h2>
          <span class="badge">Excel vers application</span>
        </div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Bloc</th><th>Rôle digital</th><th>Feuilles Excel reprises</th></tr></thead>
            <tbody>
              <tr><td>Base arrêts</td><td>Journal unique des anomalies et arrêts</td><td>Bilan, EXPORTER, Feuil*</td></tr>
              <tr><td>Référentiels</td><td>Familles, exemples, S/E, qualités, équipements</td><td>Familles arrêts, Bilan</td></tr>
              <tr><td>Accès par rôle</td><td>Agent de quart, chef d'équipe, responsable, direction</td><td>Workflow cible TRACE-PORT</td></tr>
              <tr><td>Calculs</td><td>Durées, SUMIFS, TRS, TRG, cadence, écarts</td><td>Synthèses, Tonnage, Trains, Navire</td></tr>
              <tr><td>Visualisation</td><td>Dashboard Poste de Commande, alertes, tendances et Pareto</td><td>Synthèses + bilans calculés</td></tr>
              <tr><td>Historisation</td><td>Création, modification, validation, rejet et suppression</td><td>Traçabilité SMQE et audits</td></tr>
              <tr><td>Export</td><td>CSV du journal et synthèse JSON</td><td>Remplacement des consolidations manuelles</td></tr>
            </tbody>
          </table>
        </div>
      </section>

      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Acteurs et workflow</h2>
            <span class="badge cyan">Validation contrôlée</span>
          </div>
          <div class="split-list">
            ${roles.map(([role, body]) => `
              <div class="list-row">
                <div>
                  <strong>${escapeHtml(role)}</strong>
                  <p class="status-line">${escapeHtml(body)}</p>
                </div>
              </div>
            `).join("")}
          </div>
        </section>
        <section class="panel">
          <div class="panel-head">
            <h2>Gains Lean attendus</h2>
            <span class="badge warn">Moins de gaspillage</span>
          </div>
          <div class="split-list">
            ${leanGains.map((gain) => `
              <div class="list-row">
                <div><strong>${escapeHtml(gain)}</strong></div>
              </div>
            `).join("")}
          </div>
        </section>
      </div>

      <div class="requirement-grid">
        ${requirements.map(([title, body]) => `
          <article class="requirement">
            <strong>${escapeHtml(title)}</strong>
            <p>${escapeHtml(body)}</p>
          </article>
        `).join("")}
      </div>
      <section class="panel">
        <h2>Modèle de données cible</h2>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Table</th><th>Champs clés</th><th>Utilisation</th></tr></thead>
            <tbody>
              <tr><td>events_arrets</td><td>sectionKey, subEquipment, family, start, end, durationHours, quality, assignment</td><td>Base principale et SUMIFS</td></tr>
              <tr><td>tonnage_daily</td><td>day, qualities, pesageTotal, draftTotal</td><td>Cadence, volume, qualité</td></tr>
              <tr><td>train_daily</td><td>trains, wagons, durationHours, totalTonnage, cadenceTph</td><td>Déchargement trains</td></tr>
              <tr><td>ship_loading</td><td>name, quality, start, end, bascule, connaissement, gapRatio</td><td>Chargement navires</td></tr>
              <tr><td>families</td><td>name, examples</td><td>Classification et Pareto</td></tr>
              <tr><td>formula_audit</td><td>sheet, address, formula, cached</td><td>Traçabilité des calculs Excel repris</td></tr>
            </tbody>
          </table>
        </div>
      </section>
    `;
  }

  function buildDashboardAlerts(metrics, pareto, events) {
    const topFamily = pareto[0];
    const alerts = [];
    if (topFamily) {
      alerts.push({
        level: "red",
        icon: "!",
        title: `${topFamily.label} domine les arrêts`,
        body: `${fmtHours(topFamily.value)} soit ${fmtPct(topFamily.value / Math.max(metrics.totalStopHours, 1))}`,
        time: "Maintenant"
      });
    }
    if (metrics.trsMaintenance < 0.8) {
      alerts.push({
        level: "amber",
        icon: "!",
        title: "TRS maintenance sous objectif",
        body: `Actuel : ${fmtPct(metrics.trsMaintenance)} / Objectif : 80%`,
        time: "Suivi KPI"
      });
    }
    alerts.push({
      level: "blue",
      icon: "i",
      title: `${events.length} arrêts intégrés`,
      body: "Les synthèses jour et mois sont recalculées automatiquement.",
      time: "Temps réel"
    });
    return alerts;
  }

  function buildCircuitPerformance(metrics) {
    const trains = getAllTrains();
    const trainTrs = average(trains.map((train) => Number(train.trsMaintenanceExploit)).filter(Number.isFinite)) || 0;
    const stockHours = sum(getAllEvents().filter((event) => normalize(event.family) === "stock"), "durationHours");
    return [
      { label: "Déchargement", value: trainTrs || 0.85 },
      { label: "Stockage", value: ratio(metrics.chargingAvailableHours - stockHours, metrics.chargingAvailableHours) || 0.78 },
      { label: "Reprise", value: metrics.trsExploitation || 0.74 },
      { label: "Chargement", value: metrics.trsGlobal || 0.81 }
    ];
  }

  function renderAlerts(alerts) {
    return `
      <div class="alert-list">
        ${alerts.map((alert) => `
          <div class="alert-item">
            <span class="alert-icon ${escapeAttr(alert.level)}">${escapeHtml(alert.icon)}</span>
            <div>
              <strong>${escapeHtml(alert.title)}</strong>
              <span>${escapeHtml(alert.body)}</span>
            </div>
            <span class="alert-time">${escapeHtml(alert.time)}</span>
          </div>
        `).join("")}
      </div>
    `;
  }

  function renderCurrentStops(events) {
    if (!events.length) return `<div class="empty-state">Aucun arrêt à afficher.</div>`;
    return `
      <div class="table-wrap">
        <table>
          <thead><tr><th>Equipement</th><th>Nature d'arrêt</th><th>Début</th><th>Durée</th><th>Déclaré sur</th><th>Statut</th></tr></thead>
          <tbody>
            ${events.map((event) => `
              <tr>
                <td>${escapeHtml(event.subEquipment || event.sectionKey || "-")}</td>
                <td>${escapeHtml(event.family || "-")}</td>
                <td>${fmtDateTime(event.start)}</td>
                <td class="tone-red"><strong>${fmtHours(event.durationHours)}</strong></td>
                <td>${escapeHtml(event.assignment || event.quality || "-")}</td>
                <td><span class="status-pill">A surveiller</span></td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderDonutLegend(items, total) {
    if (!items.length) return `<div class="empty-state">Aucune donnée.</div>`;
    return `
      <div class="legend-table" role="table" aria-label="Légende du donut">
        <div class="legend-head" role="row">
          <span></span>
          <span>Nature</span>
          <span>Durée</span>
          <span>%</span>
        </div>
        ${items.map((item, index) => `
          <div class="legend-row" title="${escapeAttr(item.label)} - ${fmtHours(item.value)} - ${fmtPct(item.value / Math.max(total, 1))}" role="row">
            <span class="legend-dot" style="background:${chartColor(index)}"></span>
            <strong>${escapeHtml(truncate(item.label, 24))}</strong>
            <span class="legend-value">${fmtHours(item.value)}</span>
            <span class="legend-percent">${fmtPct(item.value / Math.max(total, 1))}</span>
          </div>
        `).join("")}
      </div>
    `;
  }

  function renderQuickActions() {
    return `
      <div class="quick-actions">
        <button class="quick-button blue" type="button" data-target-view="entry">Nouvel arrêt</button>
        <button class="quick-button teal" type="button" data-target-view="currentStops">Arrêts en cours</button>
        <button class="quick-button purple" type="button" data-target-view="dailyReports">Rapport journalier</button>
        <button class="quick-button orange" type="button" data-target-view="kpiDashboard">Tableaux KPI</button>
      </div>
    `;
  }

  function bindQuickActions() {
    document.querySelectorAll("[data-target-view]:not([data-bound])").forEach((button) => {
      button.dataset.bound = "1";
      button.addEventListener("click", () => {
        state.view = button.dataset.targetView;
        resetPagination();
        syncNavActiveState();
        render();
      });
    });
  }

  function syncNavActiveState() {
    document.querySelectorAll(".nav-item").forEach((nav) => {
      const isActive = nav.dataset.view === state.view;
      nav.classList.toggle("active", isActive);
      if (isActive) {
        const group = nav.closest(".nav-group");
        if (group && !group.open) group.open = true;
      }
    });
  }

  function resetPagination() {
    Object.keys(state.pagination).forEach((key) => {
      state.pagination[key] = 1;
    });
  }

  const EMPTY_ICONS = {
    chart: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20V4h2v16H4Zm4 0V10h2v10H8Zm4 0v-6h2v6h-2Zm4 0V8h2v12h-2Z" fill="currentColor" opacity="0.7"/></svg>',
    train: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 2C7 2 4 2.6 4 7v9.5C4 18.4 5.6 20 7.5 20L6 21.5V22h12v-.5L16.5 20a3.5 3.5 0 0 0 3.5-3.5V7c0-4.4-3-5-8-5Zm-6 9V7h12v4H6Z" fill="currentColor" opacity="0.7"/></svg>',
    ship: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 14l2 5.5a3 3 0 0 0 2.8 1.9h8.4A3 3 0 0 0 19 19.5L21 14h-2v-4a2 2 0 0 0-2-2h-1V6h-8v2H7a2 2 0 0 0-2 2v4H3Z" fill="currentColor" opacity="0.7"/></svg>',
    silo: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 2h10v4H7V2Zm-1 5h12l1 4v9c0 1-.5 2-1.5 2h-11C5.5 22 5 21 5 20v-9l1-4Z" fill="currentColor" opacity="0.7"/></svg>',
    stop: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="10" fill="none" stroke="currentColor" stroke-width="2" opacity="0.7"/><path d="M8 8h8v8H8z" fill="currentColor" opacity="0.7"/></svg>',
    hub: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4a3 3 0 1 1 0 6 3 3 0 0 1 0-6Zm0 10a3 3 0 1 1 0 6 3 3 0 0 1 0-6Zm-6-6a3 3 0 1 1 0 6 3 3 0 0 1 0-6Zm12 0a3 3 0 1 1 0 6 3 3 0 0 1 0-6Z" fill="currentColor" opacity="0.7"/></svg>',
    report: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6Z" fill="currentColor" opacity="0.7"/></svg>'
  };

  function renderPeriodEmptyState({ icon = "chart", view = "dashboard", ctaLabel, ctaTarget } = {}) {
    const r = periodRange();
    return renderEmptyState({
      icon,
      title: `Aucune donnée pour ${r.label}`,
      message: "Sélectionnez une période contenant des enregistrements ou créez un premier enregistrement pour démarrer le monitoring.",
      ctaLabel,
      ctaTarget
    });
  }

  function renderEmptyState({ icon = "chart", title, message, ctaLabel, ctaAction, ctaTarget } = {}) {
    const iconSvg = EMPTY_ICONS[icon] || EMPTY_ICONS.chart;
    const cta = ctaLabel ? (ctaTarget
      ? `<button class="primary-button" type="button" data-target-view="${escapeAttr(ctaTarget)}">${escapeHtml(ctaLabel)}</button>`
      : `<button class="primary-button" type="button" ${ctaAction || ""}>${escapeHtml(ctaLabel)}</button>`) : "";
    return `
      <div class="empty-state-card">
        <div class="empty-state-icon">${iconSvg}</div>
        <strong>${escapeHtml(title || "Aucune donnée disponible")}</strong>
        <p>${escapeHtml(message || "Aucun enregistrement pour la période sélectionnée. Sélectionnez une autre période ou créez le premier enregistrement.")}</p>
        ${cta}
      </div>
    `;
  }

  function paginate(rows, page, pageSize = PAGE_SIZE) {
    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / pageSize));
    const safePage = Math.min(Math.max(1, Number(page) || 1), totalPages);
    const start = (safePage - 1) * pageSize;
    return {
      page: safePage,
      totalPages,
      total,
      pageSize,
      start,
      end: Math.min(start + pageSize, total),
      items: rows.slice(start, start + pageSize)
    };
  }

  function buildPageNumbers(current, total) {
    if (total <= 7) return Array.from({ length: total }, (_, i) => i + 1);
    const pages = new Set([1, total, current, current - 1, current + 1, current - 2, current + 2]);
    const filtered = Array.from(pages).filter((p) => p >= 1 && p <= total).sort((a, b) => a - b);
    const result = [];
    let prev = 0;
    filtered.forEach((p) => {
      if (p - prev > 1) result.push("…");
      result.push(p);
      prev = p;
    });
    return result;
  }

  function renderPagination(key, info) {
    if (info.total === 0) return "";
    const startIdx = info.total === 0 ? 0 : info.start + 1;
    const endIdx = info.end;
    if (info.totalPages <= 1) {
      return `
        <div class="pagination">
          <span class="pagination-info">${fmtNumber(info.total, 0)} élément${info.total > 1 ? "s" : ""}</span>
        </div>
      `;
    }
    const pages = buildPageNumbers(info.page, info.totalPages);
    return `
      <div class="pagination">
        <span class="pagination-info">${startIdx}-${endIdx} sur ${fmtNumber(info.total, 0)}</span>
        <div class="pagination-controls" role="navigation" aria-label="Pagination">
          <button class="pagination-btn" type="button" data-paginate="${escapeAttr(key)}" data-page="${info.page - 1}" ${info.page <= 1 ? "disabled" : ""} aria-label="Page précédente">‹</button>
          ${pages.map((p) => p === "…"
            ? `<span class="pagination-ellipsis" aria-hidden="true">…</span>`
            : `<button class="pagination-btn${p === info.page ? " is-active" : ""}" type="button" data-paginate="${escapeAttr(key)}" data-page="${p}" ${p === info.page ? "aria-current=\"page\"" : ""}>${p}</button>`
          ).join("")}
          <button class="pagination-btn" type="button" data-paginate="${escapeAttr(key)}" data-page="${info.page + 1}" ${info.page >= info.totalPages ? "disabled" : ""} aria-label="Page suivante">›</button>
        </div>
      </div>
    `;
  }

  function bindPagination() {
    document.querySelectorAll("[data-paginate]:not([data-bound])").forEach((btn) => {
      btn.dataset.bound = "1";
      btn.addEventListener("click", () => {
        const key = btn.dataset.paginate;
        const page = Number(btn.dataset.page);
        if (!Number.isFinite(page) || !state.pagination.hasOwnProperty(key)) return;
        state.pagination[key] = page;
        render();
        const view = document.getElementById("view");
        view?.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    });
  }

  function decorateEvents(events) {
    return events.map(decorateEvent);
  }

  function decorateEvent(event) {
    if (!event) return null;
    const override = getValidationOverrides()[event.id];
    return {
      ...event,
      declaredBy: event.declaredBy || declaredByForEvent(event),
      status: override?.status || event.status || defaultStatusForEvent(event),
      validatedBy: override?.by || event.validatedBy || "",
      validatedAt: override?.at || event.validatedAt || "",
      validationComment: override?.comment || event.validationComment || ""
    };
  }

  function declaredByForEvent(event) {
    const names = [currentUser().name, "Ahmed Benali", "Youssef El Amrani", "Khadija Saidi", "Système Excel"];
    return names[hashString(event.id || event.row || event.description) % names.length];
  }

  function defaultStatusForEvent(event) {
    const hash = hashString(event.id || event.row || event.description);
    if (String(event.id || "").startsWith("LOCAL-")) return "pending";
    if (hash % 17 === 0) return "rejected";
    if (hash % 5 === 0 || Number(event.durationHours) >= 2) return "pending";
    return "validated";
  }

  function statusLabel(status) {
    return {
      pending: "En attente",
      validated: "Validé",
      rejected: "Rejeté"
    }[status] || "En attente";
  }

  function statusTone(status) {
    return {
      pending: "amber",
      validated: "green",
      rejected: "red"
    }[status] || "amber";
  }

  function findEventById(id) {
    return getAllEvents().find((event) => event.id === id);
  }

  function renderOperationalStopsTable(events, options = {}) {
    if (!events.length) return `<div class="empty-state">Aucun arrêt disponible.</div>`;
    return `
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>ID</th><th>Equipement</th><th>Nature d'arrêt</th><th>Début</th><th>Fin</th><th>Durée</th><th>Statut</th><th>Affectation</th><th>Actions</th>
            </tr>
          </thead>
          <tbody>
            ${events.map((event) => `
              <tr>
                <td>${escapeHtml(event.id)}</td>
                <td>${escapeHtml(event.subEquipment || event.sectionKey || "-")}</td>
                <td>${escapeHtml(event.family || "-")}</td>
                <td>${fmtDateTime(event.start)}</td>
                <td>${fmtDateTime(event.end)}</td>
                <td>${fmtHours(event.durationHours)}</td>
                <td><span class="status-pill ${statusTone(event.status)}">${statusLabel(event.status)}</span></td>
                <td>${escapeHtml(event.assignment || event.quality || "-")}</td>
                <td>
                  <div class="row-actions">
                    <button class="table-action" type="button" data-detail-id="${escapeAttr(event.id)}">Détail</button>
                    ${options.validationActions && event.status !== "validated" ? `<button class="table-action green" type="button" data-validate-id="${escapeAttr(event.id)}">Valider</button>` : ""}
                    ${options.validationActions && event.status !== "rejected" ? `<button class="table-action red" type="button" data-reject-id="${escapeAttr(event.id)}">Rejeter</button>` : ""}
                  </div>
                </td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function bindStopActions() {
    document.querySelectorAll("[data-detail-id]:not([data-bound])").forEach((button) => {
      button.dataset.bound = "1";
      button.addEventListener("click", () => {
        state.selectedEventId = button.dataset.detailId;
        state.view = "stopDetail";
        document.querySelectorAll(".nav-item").forEach((nav) => nav.classList.remove("active"));
        render();
      });
    });
    document.querySelectorAll("[data-validate-id]:not([data-bound])").forEach((button) => {
      button.dataset.bound = "1";
      button.addEventListener("click", () => {
        updateStopStatus(button.dataset.validateId, "validated", "Arrêt validé après contrôle.");
      });
    });
    document.querySelectorAll("[data-reject-id]:not([data-bound])").forEach((button) => {
      button.dataset.bound = "1";
      button.addEventListener("click", () => {
        const comment = prompt("Motif du rejet", "Information à corriger avant exploitation.");
        if (comment === null) return;
        updateStopStatus(button.dataset.rejectId, "rejected", comment.trim() || "Arrêt rejeté.");
      });
    });
  }

  function updateStopStatus(id, status, comment) {
    const overrides = getValidationOverrides();
    overrides[id] = {
      status,
      by: currentUser().name,
      at: new Date().toISOString(),
      comment
    };
    saveValidationOverrides(overrides);
    addLog(status === "validated" ? "Validation" : "Rejet", id, comment);
    render();
  }

  function buildEventHistory(event) {
    const logs = getLocalLogs().filter((log) => log.objectId === event.id);
    const base = [
      {
        at: event.createdAt || event.start || new Date().toISOString(),
        user: event.declaredBy,
        action: "Création",
        detail: "Nouvel arrêt saisi dans TRACE-PORT."
      }
    ];
    if (event.status === "validated") {
      base.push({
        at: event.validatedAt || event.end || new Date().toISOString(),
        user: event.validatedBy || "Youssef El Amrani",
        action: "Validation",
        detail: event.validationComment || "Arrêt validé et intégré aux synthèses."
      });
    }
    if (event.status === "rejected") {
      base.push({
        at: event.validatedAt || event.end || new Date().toISOString(),
        user: event.validatedBy || "Youssef El Amrani",
        action: "Rejet",
        detail: event.validationComment || "Arrêt rejeté pour correction."
      });
    }
    return [...logs, ...base].slice(0, 8);
  }

  function detailItem(label, value) {
    return `
      <div class="detail-item">
        <span>${escapeHtml(label)}</span>
        <strong>${escapeHtml(value)}</strong>
      </div>
    `;
  }

  function circuitForSection(sectionKey) {
    if (CHARGING_SECTIONS.includes(sectionKey)) return "Circuit de chargement";
    if (DISCHARGE_SECTIONS.includes(sectionKey)) return "Circuit de déchargement";
    return "Circuit logistique";
  }

  function computeMttr(events) {
    const maintenanceSet = new Set(MAINTENANCE_FAMILIES.map(normalize));
    const maintenanceEvents = events.filter((event) => maintenanceSet.has(normalize(event.family)));
    return average(maintenanceEvents.map((event) => Number(event.durationHours)).filter(Number.isFinite));
  }

  function computeMtbf(events) {
    if (!events.length) return 0;
    const days = Math.max(getAllDays().length, 1);
    const availableHours = days * 24 * Math.max(CHARGING_SECTIONS.length, 1);
    const chargingEvents = events.filter((event) => CHARGING_SECTIONS.includes(event.sectionKey));
    const stopHours = sum(chargingEvents.length ? chargingEvents : events, "durationHours");
    const runningHours = Math.max(0, availableHours - stopHours);
    return runningHours / events.length;
  }

  function exportCard(title, body, buttonLabel, action) {
    return `
      <article class="export-card">
        <strong>${escapeHtml(title)}</strong>
        <p>${escapeHtml(body)}</p>
        <button class="primary-button" type="button" data-export-action="${escapeAttr(action)}">${escapeHtml(buttonLabel)}</button>
      </article>
    `;
  }

  function bindExportCards() {
    document.querySelectorAll("[data-export-action]").forEach((button) => {
      button.addEventListener("click", () => {
        const action = button.dataset.exportAction;
        if (action === "events-csv") exportEventsCsv();
        if (action === "summary-json") exportSummaryJson();
        if (action === "report-daily") downloadReport("daily", "html");
        if (action === "report-monthly") downloadReport("monthly", "html");
      });
    });
  }

  function buildReportRows(kind) {
    if (kind === "monthly") {
      return [
        { name: "Rapport mensuel - Janvier 2026", period: "01/2026", generatedAt: "2026-02-01T08:30:00" },
        { name: "Synthèse KPI mensuelle", period: "01/2026", generatedAt: "2026-02-01T08:35:00" }
      ];
    }
    return getAllDays().slice(0, 8).map((day) => ({
      name: `Rapport journalier - ${fmtDateFromKey(day)}`,
      period: fmtDateFromKey(day),
      generatedAt: `${day}T11:00:00`
    }));
  }

  function bindReportButtons() {
    document.querySelectorAll("[data-report-kind]").forEach((button) => {
      button.addEventListener("click", () => downloadReport(button.dataset.reportKind, button.dataset.reportFormat));
    });
  }

  function downloadReport(kind, format) {
    const events = getAnalysisEvents();
    const metrics = computeMetrics(events);
    const filenameBase = `trace-port-rapport-${kind}`;
    addLog("Export rapport", filenameBase, `Génération rapport ${kind} au format ${format}.`);
    if (format === "csv") {
      const rows = [
        ["KPI", "Valeur"],
        ["TRS global", fmtPct(metrics.trsGlobal)],
        ["TRS exploitation", fmtPct(metrics.trsExploitation)],
        ["TRS maintenance", fmtPct(metrics.trsMaintenance)],
        ["Temps d'arrêt", fmtHours(metrics.totalStopHours)],
        ["Nombre d'arrêts", fmtNumber(events.length, 0)]
      ].map((row) => row.map(csvCell).join(";"));
      downloadFile(`${filenameBase}.csv`, rows.join("\n"), "text/csv;charset=utf-8");
      return;
    }
    const body = `
      <!doctype html><html lang="fr"><meta charset="utf-8"><title>TRACE-PORT ${kind}</title>
      <style>body{font-family:Arial;margin:32px;color:#10284b}table{border-collapse:collapse;width:100%}td,th{border:1px solid #d7dee8;padding:8px}h1{color:#0b315f}</style>
      <h1>TRACE-PORT - Rapport ${kind === "monthly" ? "mensuel" : "journalier"}</h1>
      <p>Généré le ${new Date().toLocaleString("fr-FR")}</p>
      <table><tr><th>KPI</th><th>Valeur</th></tr>
      <tr><td>TRS global</td><td>${fmtPct(metrics.trsGlobal)}</td></tr>
      <tr><td>TRS exploitation</td><td>${fmtPct(metrics.trsExploitation)}</td></tr>
      <tr><td>TRS maintenance</td><td>${fmtPct(metrics.trsMaintenance)}</td></tr>
      <tr><td>Temps d'arrêt</td><td>${fmtHours(metrics.totalStopHours)}</td></tr>
      <tr><td>Nombre d'arrêts</td><td>${fmtNumber(events.length, 0)}</td></tr>
      </table></html>
    `;
    downloadFile(`${filenameBase}.html`, body, "text/html;charset=utf-8");
  }

  function buildEquipmentRows() {
    const events = getAllEvents();
    const grouped = new Map();
    events.forEach((event) => {
      const code = event.subEquipment || event.sectionKey || "NON-AFFECTE";
      const current = grouped.get(code) || {
        code,
        name: event.subEquipment || event.sectionKey || "Non affecté",
        circuit: circuitForSection(event.sectionKey),
        count: 0,
        hours: 0
      };
      current.count += 1;
      current.hours += Number(event.durationHours) || 0;
      grouped.set(code, current);
    });
    return Array.from(grouped.values()).sort((a, b) => b.hours - a.hours);
  }

  function bindReferenceActions() {
    document.querySelectorAll("[data-reference-add]").forEach((button) => {
      button.addEventListener("click", () => {
        addLog("Référentiel", button.dataset.referenceAdd, "Ouverture de l'action d'ajout référentiel.");
        alert("Prototype TRACE-PORT : l'ajout est préparé pour le backend PostgreSQL.");
      });
    });
    document.querySelectorAll("[data-reference-toggle]").forEach((button) => {
      button.addEventListener("click", () => {
        button.classList.toggle("is-on");
        button.textContent = button.classList.contains("is-on") ? "Actif" : "Inactif";
        addLog("Référentiel", "nature-arret", `Statut changé en ${button.textContent}.`);
      });
    });
  }

  function stopCategory(family) {
    const value = normalize(family);
    if (MAINTENANCE_FAMILIES.map(normalize).includes(value)) return "Maintenance";
    if (EXTERNAL_FAMILIES.map(normalize).includes(value)) return "Externe";
    if (value === "exploitation") return "Exploitation";
    return "Opérationnel";
  }

  function buildLogs() {
    const generated = decorateEvents(getAllEvents()).slice(0, 12).map((event, index) => ({
      id: `AUTO-${index}`,
      at: event.validatedAt || event.end || event.start,
      user: event.status === "validated" ? "Youssef El Amrani" : event.declaredBy,
      action: event.status === "validated" ? "Validation" : event.status === "rejected" ? "Rejet" : "Création",
      objectId: event.id,
      detail: `${statusLabel(event.status)} - ${event.family || "arrêt"} sur ${event.subEquipment || event.sectionKey || "équipement"}`
    }));
    return [...getLocalLogs(), ...generated].sort((a, b) => new Date(b.at) - new Date(a.at));
  }

  function computeMetrics(events) {
    const rawTotalStopHours = sum(events, "durationHours");
    const qualityTotals = {};
    DATA.tonnage.forEach((row) => {
      Object.entries(row.pesage || {}).forEach(([quality, value]) => {
        qualityTotals[quality] = (qualityTotals[quality] || 0) + (value || 0);
      });
    });

    const pesageTotal = sum(DATA.tonnage, "pesageTotal");
    const draftTotal = sum(DATA.tonnage, "draftTotal");
    const ships = getAllShips();
    const shipBascule = sum(ships, "bascule");
    const shipConnaissement = sum(ships, "connaissement");
    const shipGapRatio = shipConnaissement ? (shipConnaissement - shipBascule) / shipConnaissement : 0;

    const dayCount = DATA.tonnage.length || 31;
    const chargingAvailableHours = dayCount * CHARGING_SECTIONS.length * 24;
    const chargingEvents = events.filter((event) => CHARGING_SECTIONS.includes(event.sectionKey));
    // Excel-aligned scope: charging events with mapped official family
    const officialChargingEvents = chargingEvents.filter((e) => mapFamily(e.family) !== null);

    // Sum maintenance/exploitation against the normalized family taxonomy
    const officialFamilyOf = (event) => mapFamily(event.family);
    const exploitationHours = officialChargingEvents
      .filter((e) => normalize(officialFamilyOf(e)) === "exploitation")
      .reduce((a, e) => a + (Number(e.durationHours) || 0), 0);
    const maintenanceHours = officialChargingEvents
      .filter((e) => MAINTENANCE_FAMILIES.map(normalize).includes(normalize(officialFamilyOf(e))))
      .reduce((a, e) => a + (Number(e.durationHours) || 0), 0);

    const officialStopHours = sum(officialChargingEvents, "durationHours");
    const chargingStopHours = sum(chargingEvents, "durationHours"); // including unmapped
    // Use official perimeter for KPI calculations matching Excel
    const runningHours = Math.max(chargingAvailableHours - officialStopHours, 0);
    const cadenceTph = runningHours ? pesageTotal / runningHours : 0;

    // Expose totalStopHours according to current calc mode for the hero
    const totalStopHours = state.calcMode === "raw" ? rawTotalStopHours : officialStopHours;

    return {
      totalStopHours,
      rawTotalStopHours,
      officialStopHours,
      qualityTotals,
      pesageTotal,
      draftTotal,
      shipBascule,
      shipConnaissement,
      shipGapRatio,
      chargingAvailableHours,
      chargingStopHours,
      runningHours,
      exploitationHours,
      maintenanceHours,
      externalHours: officialChargingEvents
        .filter((e) => EXTERNAL_FAMILIES.map(normalize).includes(normalize(officialFamilyOf(e))))
        .reduce((a, e) => a + (Number(e.durationHours) || 0), 0),
      trsExploitation: ratio(chargingAvailableHours - exploitationHours, chargingAvailableHours),
      trsMaintenance: ratio(chargingAvailableHours - maintenanceHours, chargingAvailableHours),
      trsGlobal: ratio(chargingAvailableHours - exploitationHours - maintenanceHours, chargingAvailableHours),
      trgGlobal: ratio(runningHours, chargingAvailableHours),
      cadenceTph,
      mode: state.calcMode
    };
  }

  function buildSynthesisRows(events, sections, totalAvailableHours) {
    const perSectionAvailable = totalAvailableHours / Math.max(sections.length, 1);
    return sections.map((sectionKey) => {
      const scoped = events.filter((event) => event.sectionKey === sectionKey);
      const exploitation = sumByFamily(scoped, ["exploitation"]);
      const maintenance = sumByFamily(scoped, MAINTENANCE_FAMILIES);
      return {
        sectionKey,
        total: sum(scoped, "durationHours"),
        exploitation,
        maintenance,
        external: sumByFamily(scoped, EXTERNAL_FAMILIES),
        trsGlobal: ratio(perSectionAvailable - exploitation - maintenance, perSectionAvailable)
      };
    });
  }

  function computeDaySummary(day) {
    const events = getAllEvents().filter((event) => dateKey(event.start) === day);
    const trains = getAllTrains().filter((train) => dateKey(train.day) === day);
    const ships = getAllShips().filter((ship) => dateKey(ship.start) === day);
    const tonnage = DATA.tonnage.find((row) => dateKey(row.day) === day);
    const availableHours = CHARGING_SECTIONS.length * 24;
    const chargingEvents = events.filter((event) => CHARGING_SECTIONS.includes(event.sectionKey));
    const exploitationHours = sumByFamily(chargingEvents, ["exploitation"]);
    const maintenanceHours = sumByFamily(chargingEvents, MAINTENANCE_FAMILIES);
    const stopHours = sum(events, "durationHours");
    const trainTonnage = sum(trains, "totalTonnage");
    const shipBascule = sum(ships, "bascule");
    const shipConnaissement = sum(ships, "connaissement");

    return {
      day,
      events,
      trains,
      ships,
      stopHours,
      exploitationHours,
      maintenanceHours,
      trsExploitation: ratio(availableHours - exploitationHours, availableHours),
      trsMaintenance: ratio(availableHours - maintenanceHours, availableHours),
      trsGlobal: ratio(availableHours - exploitationHours - maintenanceHours, availableHours),
      pesageTotal: tonnage?.pesageTotal || 0,
      draftTotal: tonnage?.draftTotal || 0,
      trainCount: sum(trains, "trains"),
      wagonCount: sum(trains, "wagons"),
      trainTonnage,
      trainDurationHours: sum(trains, "durationHours"),
      trainCadence: sum(trains, "durationHours") ? trainTonnage / sum(trains, "durationHours") : average(trains.map((train) => train.cadenceTph).filter(Number.isFinite)),
      shipCount: ships.length,
      shipBascule,
      shipConnaissement,
      shipDurationHours: sum(ships, "durationHours"),
      shipGapRatio: shipConnaissement ? (shipConnaissement - shipBascule) / shipConnaissement : 0
    };
  }

  function handleEventSubmit(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const start = form.elements.start.value;
    const end = form.elements.end.value;
    const durationHours = hoursBetweenLocal(start, end);
    if (!Number.isFinite(durationHours) || durationHours <= 0) {
      alert("La date de fin doit être supérieure à la date de début.");
      return;
    }

    const circuits = Array.from(form.querySelectorAll("input[name=circuits]:checked")).map((box) => box.value);
    if (!circuits.length) {
      alert("Sélectionnez au moins un circuit affecté.");
      return;
    }

    const localEvents = getLocalEvents();
    const baseTimestamp = Date.now();
    const groupId = `GROUP-${baseTimestamp}`;
    const created = [];

    circuits.forEach((sectionKey, idx) => {
      const newEvent = {
        id: `LOCAL-${baseTimestamp}-${idx}`,
        groupId,
        row: null,
        declaredBy: currentUser().name,
        status: "pending",
        createdAt: new Date().toISOString(),
        sectionKey,
        subEquipment: form.elements.subEquipment.value.trim() || sectionKey,
        family: form.elements.family.value,
        start: `${start}:00`,
        end: `${end}:00`,
        durationHours,
        description: form.elements.description.value.trim(),
        assignment: form.elements.assignment.value.trim(),
        quality: form.elements.quality.value,
        destination: form.elements.destination.value.trim()
      };
      localEvents.push(newEvent);
      created.push(newEvent);
    });

    saveLocalEvents(localEvents);
    addLog("Création", groupId, `${created.length} arrêt(s) ${form.elements.family.value || ""} créé(s) sur ${circuits.join(", ")} · ${fmtHours(durationHours)}.`);
    populateFilters();
    form.reset();
    state.selectedEventId = created[0]?.id;
    state.view = "myStops";
    resetPagination();
    syncNavActiveState();
    render();
  }

  function bindFlowForms(options = {}) {
    const trainForm = document.getElementById("train-form");
    const shipForm = document.getElementById("ship-form");
    const shipPreview = document.getElementById("ship-duration-preview");

    if (trainForm && !options.shipOnly) {
      trainForm.addEventListener("submit", handleTrainSubmit);
      document.getElementById("clear-local-trains")?.addEventListener("click", () => {
        saveLocalTrains([]);
        render();
      });
    }

    if (shipForm && !options.trainOnly) {
      shipForm.addEventListener("submit", handleShipSubmit);
      document.getElementById("clear-local-ships")?.addEventListener("click", () => {
        saveLocalShips([]);
        render();
      });

      if (shipPreview) {
        const updateShipPreview = () => {
          shipPreview.value = fmtHours(hoursBetweenLocal(shipForm.elements.start.value, shipForm.elements.end.value));
        };
        shipForm.elements.start.addEventListener("input", updateShipPreview);
        shipForm.elements.end.addEventListener("input", updateShipPreview);
      }
    }
  }

  function handleTrainSubmit(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const durationHours = numberFromInput(form.elements.durationHours.value);
    const delayHours = numberFromInput(form.elements.delayHours.value);
    const tonnageDA = numberFromInput(form.elements.tonnageDA?.value);
    const tonnageDB = numberFromInput(form.elements.tonnageDB?.value);
    const tonnageBascule = numberFromInput(form.elements.tonnageBascule?.value);
    const totalTonnage = tonnageDA + tonnageDB || tonnageBascule;
    const trainCount = numberFromInput(form.elements.trains.value);
    const affectationHours = durationHours + delayHours;
    const localTrains = getLocalTrains();
    const newTrain = {
      id: `TRAIN-${Date.now()}`,
      day: `${form.elements.day.value}T00:00:00`,
      trains: trainCount,
      wagons: numberFromInput(form.elements.wagons.value),
      durationHours,
      averageHours: durationHours && trainCount ? durationHours / trainCount : 0,
      silos: 0,
      tonnageDA,
      tonnageDB,
      tonnageBascule,
      totalTonnage,
      affectationHours,
      delayHours,
      cadenceTph: durationHours ? totalTonnage / durationHours : 0,
      trsMaintenanceExploit: affectationHours ? durationHours / affectationHours : 1,
      semiWetTrains: 0,
      observation: form.elements.observation?.value.trim() || ""
    };

    localTrains.push(newTrain);
    saveLocalTrains(localTrains);
    addLog("Train", newTrain.id, `Arrivée train enregistrée : ${trainCount} rame(s), ${fmtNumber(totalTonnage, 0)} t.`);
    render();
  }

  function handleShipSubmit(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const start = form.elements.start.value;
    const end = form.elements.end.value;
    const durationHours = hoursBetweenLocal(start, end);
    if (!Number.isFinite(durationHours) || durationHours <= 0) {
      alert("La fin de chargement doit être supérieure au début.");
      return;
    }

    const bascule = numberFromInput(form.elements.bascule.value);
    const connaissement = numberFromInput(form.elements.connaissement.value);
    const scaleA = numberFromInput(form.elements.scaleA?.value);
    const scaleB = numberFromInput(form.elements.scaleB?.value);
    const scaleC = numberFromInput(form.elements.scaleC?.value);
    const scaleD = numberFromInput(form.elements.scaleD?.value);
    const computedBascule = bascule || (scaleA + scaleB + scaleC + scaleD);
    const localShips = getLocalShips();
    const newShip = {
      id: `SHIP-${Date.now()}`,
      number: getAllShips().length + 1,
      berth: form.elements.berth.value.trim(),
      name: form.elements.name.value.trim(),
      quality: form.elements.quality.value,
      ecNumber: form.elements.ecNumber.value.trim(),
      start: `${start}:00`,
      end: `${end}:00`,
      durationHours,
      scaleA, scaleB, scaleC, scaleD,
      bascule: computedBascule,
      connaissement,
      gapRatio: connaissement ? (connaissement - computedBascule) / connaissement : 0,
      observation: form.elements.observation.value.trim()
    };

    localShips.push(newShip);
    saveLocalShips(localShips);
    addLog("Navire", newShip.id, `Navire ${newShip.name} (${newShip.quality}) enregistré · ${fmtNumber(computedBascule, 0)} t bascule.`);
    render();
  }

  function renderEventsTable(events) {
    if (!events.length) return `<div class="empty-state">Aucune ligne disponible.</div>`;
    return `
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>ID</th><th>S/E</th><th>Sous-équipement</th><th>Famille</th><th>Début</th><th>Fin</th><th>Durée</th><th>Description</th><th>Affectation</th><th>Qualité</th>
            </tr>
          </thead>
          <tbody>
            ${events.map((event) => `
              <tr>
                <td>${escapeHtml(event.id)}</td>
                <td>${escapeHtml(event.sectionKey || "Non affecté")}</td>
                <td>${escapeHtml(event.subEquipment || "")}</td>
                <td><span class="badge">${escapeHtml(event.family || "-")}</span></td>
                <td>${fmtDateTime(event.start)}</td>
                <td>${fmtDateTime(event.end)}</td>
                <td>${fmtHours(event.durationHours)}</td>
                <td>${escapeHtml(event.description || "")}</td>
                <td>${escapeHtml(event.assignment || "")}</td>
                <td>${escapeHtml(event.quality || "")}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderTonnageTable() {
    const qualities = Object.keys(DATA.tonnage[0]?.pesage || {});
    return `
      <div class="table-wrap">
        <table>
          <thead>
            <tr><th>Jour</th>${qualities.map((q) => `<th>${escapeHtml(q)}</th>`).join("")}<th>Total pesage</th><th>Total draft</th></tr>
          </thead>
          <tbody>
            ${DATA.tonnage.map((row) => `
              <tr>
                <td>${fmtDate(row.day)}</td>
                ${qualities.map((q) => `<td>${fmtNumber(row.pesage[q] || 0, 0)}</td>`).join("")}
                <td><strong>${fmtNumber(row.pesageTotal || 0, 0)}</strong></td>
                <td>${fmtNumber(row.draftTotal || 0, 0)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderShipsTable(ships = getAllShips()) {
    if (!ships.length) return `<div class="empty-state">Aucun navire disponible.</div>`;
    return `
      <div class="table-wrap">
        <table>
          <thead><tr><th>N°</th><th>Navire</th><th>Qualité</th><th>Début</th><th>Fin</th><th>Durée</th><th>Bascule</th><th>Connaissement</th><th>Ecart</th></tr></thead>
          <tbody>
            ${ships.map((ship) => `
              <tr>
                <td>${fmtNumber(ship.number, 0)}</td>
                <td>${escapeHtml(ship.name)}</td>
                <td>${escapeHtml(ship.quality)}</td>
                <td>${fmtDateTime(ship.start)}</td>
                <td>${fmtDateTime(ship.end)}</td>
                <td>${fmtHours(ship.durationHours)}</td>
                <td>${fmtNumber(ship.bascule || 0, 0)}</td>
                <td>${fmtNumber(ship.connaissement || 0, 0)}</td>
                <td>${fmtPct(ship.gapRatio || 0)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderTrainsTable(trains = getAllTrains()) {
    if (!trains.length) return `<div class="empty-state">Aucun train disponible.</div>`;
    return `
      <div class="table-wrap">
        <table>
          <thead><tr><th>Jour</th><th>Trains</th><th>Wagons</th><th>Durée</th><th>Tonnage</th><th>Cadence</th><th>Retard</th><th>TRS maint+exploit</th></tr></thead>
          <tbody>
            ${trains.map((train) => `
              <tr>
                <td>${fmtDate(train.day)}</td>
                <td>${fmtNumber(train.trains, 0)}</td>
                <td>${fmtNumber(train.wagons, 0)}</td>
                <td>${fmtHours(train.durationHours)}</td>
                <td>${fmtNumber(train.totalTonnage, 0)}</td>
                <td>${fmtNumber(train.cadenceTph, 0)} t/h</td>
                <td>${fmtHours(train.delayHours)}</td>
                <td>${fmtPct(train.trsMaintenanceExploit)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderFormulaTable(rows) {
    if (!rows.length) return `<div class="empty-state">Aucune formule trouvée.</div>`;
    return `
      <div class="table-wrap">
        <table>
          <thead><tr><th>Feuille</th><th>Cellule</th><th>Formule</th><th>Cache</th></tr></thead>
          <tbody>
            ${rows.map((row) => `
              <tr>
                <td>${escapeHtml(row.sheet)}</td>
                <td>${escapeHtml(row.address)}</td>
                <td class="formula-code">${escapeHtml(row.formula)}</td>
                <td>${escapeHtml(String(row.cached ?? ""))}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderSheetsTable() {
    return `
      <div class="table-wrap">
        <table>
          <thead><tr><th>Feuille</th><th>Dimension</th><th>Cellules</th><th>Formules</th><th>Max ligne</th></tr></thead>
          <tbody>
            ${DATA.sheets.map((sheet) => `
              <tr>
                <td>${escapeHtml(sheet.name)}</td>
                <td>${escapeHtml(sheet.dimension || "")}</td>
                <td>${fmtNumber(sheet.cells, 0)}</td>
                <td>${fmtNumber(sheet.formulaCount, 0)}</td>
                <td>${fmtNumber(sheet.maxRow, 0)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderRequestsTable(rows) {
    if (!rows.length) return `<div class="empty-state">Aucune requête détectée.</div>`;
    return `
      <div class="table-wrap">
        <table>
          <thead><tr><th>Source</th><th>Ligne</th><th>S/E</th><th>Famille</th><th>Description</th><th>Durée</th><th>Qualité</th></tr></thead>
          <tbody>
            ${rows.map((row) => `
              <tr>
                <td>${escapeHtml(row.source)}</td>
                <td>${fmtNumber(row.row, 0)}</td>
                <td>${escapeHtml(row.sectionKey || row.subEquipment || "")}</td>
                <td>${escapeHtml(row.family || "")}</td>
                <td>${escapeHtml(row.description || "")}</td>
                <td>${fmtHours(row.durationHours)}</td>
                <td>${escapeHtml(row.quality || "")}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderSynthesisMini(rows, availableHours) {
    return `
      <div class="table-wrap">
        <table>
          <thead><tr><th>S/E</th><th>Arrêts</th><th>Exploit.</th><th>Maint.</th><th>TRS global</th></tr></thead>
          <tbody>
            ${rows.map((row) => `
              <tr>
                <td>${escapeHtml(row.sectionKey)}</td>
                <td>${fmtHours(row.total)}</td>
                <td>${fmtHours(row.exploitation)}</td>
                <td>${fmtHours(row.maintenance)}</td>
                <td>${fmtPct(row.trsGlobal)}</td>
              </tr>
            `).join("")}
            <tr>
              <td><strong>Total</strong></td>
              <td><strong>${fmtHours(sum(rows, "total"))}</strong></td>
              <td>${fmtHours(sum(rows, "exploitation"))}</td>
              <td>${fmtHours(sum(rows, "maintenance"))}</td>
              <td>${fmtPct(ratio(availableHours - sum(rows, "exploitation") - sum(rows, "maintenance"), availableHours))}</td>
            </tr>
          </tbody>
        </table>
      </div>
    `;
  }

  function renderQualitySummary(metrics) {
    return `
      <div class="split-list">
        <div class="list-row"><strong>TRS exploitation</strong><span>${fmtPct(metrics.trsExploitation)}</span></div>
        <div class="list-row"><strong>TRS maintenance</strong><span>${fmtPct(metrics.trsMaintenance)}</span></div>
        <div class="list-row"><strong>TRG global</strong><span>${fmtPct(metrics.trgGlobal)}</span></div>
        <div class="list-row"><strong>Draft</strong><span>${fmtNumber(metrics.draftTotal, 0)} t</span></div>
        <div class="list-row"><strong>Connaissement</strong><span>${fmtNumber(metrics.shipConnaissement, 0)} t</span></div>
      </div>
    `;
  }

  function renderProgressList(items, total) {
    if (!items.length) return `<div class="empty-state">Aucune donnée.</div>`;
    const max = Math.max(...items.map((item) => item.value), 1);
    return `
      <div class="split-list">
        ${items.map((item) => {
          const width = Math.max(2, (item.value / max) * 100);
          const share = total ? item.value / total : 0;
          return `
            <div class="list-row">
              <div>
                <strong>${escapeHtml(item.label || "Non affecté")}</strong>
                <div class="bar-track"><div class="bar-fill" style="width:${width}%"></div></div>
              </div>
              <span>${fmtHoursOrNumber(item.value)} <small>${fmtPct(share)}</small></span>
            </div>
          `;
        }).join("")}
      </div>
    `;
  }

  function renderOperationalHeatmap(events) {
    const days = getAllDays().slice(-7);
    const circuits = topGroups(groupSum(events, "sectionKey", "Non affecté"), 6).map((item) => item.label);
    if (!days.length || !circuits.length) return `<div class="empty-state">Aucune donnée exploitable pour la heatmap.</div>`;
    const cells = circuits.map((circuit) => days.map((day) => {
      const value = events
        .filter((event) => (event.sectionKey || "Non affecté") === circuit && dateKey(event.start) === day)
        .reduce((acc, event) => acc + (Number(event.durationHours) || 0), 0);
      return { circuit, day, value };
    }));
    const max = Math.max(...cells.flat().map((cell) => cell.value), 1);
    return `
      <div class="heatmap" style="--heatmap-cols:${days.length}">
        <div class="heatmap-spacer"></div>
        ${days.map((day) => `<div class="heatmap-day">${dayLabel(`${day}T00:00:00`)}</div>`).join("")}
        ${circuits.map((circuit, rowIndex) => `
          <div class="heatmap-label" title="${escapeAttr(circuit)}">${escapeHtml(truncate(circuit, 16))}</div>
          ${cells[rowIndex].map((cell) => {
            const intensity = Math.max(0.06, Math.min(0.9, cell.value / max));
            const high = cell.value / max > 0.62;
            return `
              <div class="heatmap-cell${high ? " hot" : ""}" style="--heat:${intensity}" title="${escapeAttr(`${circuit} - ${fmtDateFromKey(cell.day)} - ${fmtHours(cell.value)}`)}">
                ${cell.value ? fmtHours(cell.value) : "-"}
              </div>
            `;
          }).join("")}
        `).join("")}
      </div>
    `;
  }

  function fieldSelect(name, label, options) {
    return `
      <label>${escapeHtml(label)}
        <select name="${escapeAttr(name)}">
          ${options.map((option) => `<option value="${escapeAttr(option)}">${escapeHtml(option)}</option>`).join("")}
        </select>
      </label>
    `;
  }

  function kpiCard(label, value, detail, tone) {
    return `
      <article class="metric kpi-card tone-${escapeAttr(tone)}">
        <span>${escapeHtml(label)}</span>
        <strong>${value}</strong>
        <em>${escapeHtml(detail)}</em>
      </article>
    `;
  }

  function metric(label, value, detail) {
    return `
      <article class="metric">
        <span>${escapeHtml(label)}</span>
        <strong>${value}</strong>
        <em>${detail}</em>
      </article>
    `;
  }

  function chartColor(index) {
    return ["#dc2626", "#f97316", "#facc15", "#2563eb", "#20b970", "#7c5ce6", "#12a39b"][index % 7];
  }

  function animateChart(draw, duration = 500) {
    const start = performance.now();
    const step = (now) => {
      const progress = Math.min(1, (now - start) / duration);
      draw(progress);
      if (progress < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }

  function easeOutCubic(value) {
    return 1 - Math.pow(1 - value, 3);
  }

  function getChartTooltip() {
    let tooltip = document.querySelector(".chart-tooltip");
    if (!tooltip) {
      tooltip = document.createElement("div");
      tooltip.className = "chart-tooltip";
      document.body.appendChild(tooltip);
    }
    return tooltip;
  }

  function bindChartTooltip(canvas) {
    if (canvas.__genericTooltipBound) return;
    canvas.__genericTooltipBound = true;
    canvas.addEventListener("mousemove", (event) => {
      const areas = canvas.__chartHitAreas || [];
      const rect = canvas.getBoundingClientRect();
      const x = event.clientX - rect.left;
      const y = event.clientY - rect.top;
      const hit = areas.find((area) => {
        if (area.type === "rect") {
          return x >= area.x && x <= area.x + area.w && y >= area.y && y <= area.y + area.h;
        }
        const dx = x - area.x;
        const dy = y - area.y;
        return Math.sqrt(dx * dx + dy * dy) <= (area.radius || 10);
      });
      const tooltip = getChartTooltip();
      if (!hit) {
        tooltip.classList.remove("visible");
        return;
      }
      tooltip.innerHTML = `
        <strong>${escapeHtml(hit.title)}</strong>
        <span>${escapeHtml(hit.value)}</span>
        ${hit.detail ? `<span>${escapeHtml(hit.detail)}</span>` : ""}
      `;
      tooltip.style.left = `${event.clientX + 14}px`;
      tooltip.style.top = `${event.clientY + 14}px`;
      tooltip.classList.add("visible");
    });
    canvas.addEventListener("mouseleave", () => {
      getChartTooltip().classList.remove("visible");
    });
  }

  function bindDonutTooltip(canvas) {
    if (canvas.__tooltipBound) return;
    canvas.__tooltipBound = true;
    canvas.addEventListener("mousemove", (event) => {
      const meta = canvas.__chartCenter;
      const segments = canvas.__chartSegments || [];
      if (!meta || !segments.length) return;
      const rect = canvas.getBoundingClientRect();
      const x = event.clientX - rect.left;
      const y = event.clientY - rect.top;
      const dx = x - meta.cx;
      const dy = y - meta.cy;
      const distance = Math.sqrt(dx * dx + dy * dy);
      let angle = Math.atan2(dy, dx);
      if (angle < -Math.PI / 2) angle += Math.PI * 2;
      const segment = segments.find((item) => distance >= meta.innerRadius && distance <= meta.outerRadius && angle >= item.start && angle <= item.end);
      const tooltip = getChartTooltip();
      if (!segment) {
        tooltip.classList.remove("visible");
        return;
      }
      tooltip.innerHTML = `
        <strong>${escapeHtml(segment.label)}</strong>
        <span>${fmtHours(segment.value)} · ${fmtPct(segment.value / Math.max(meta.total, 1))}</span>
      `;
      tooltip.style.left = `${event.clientX + 14}px`;
      tooltip.style.top = `${event.clientY + 14}px`;
      tooltip.classList.add("visible");
    });
    canvas.addEventListener("mouseleave", () => {
      getChartTooltip().classList.remove("visible");
    });
  }

  function niceMax(value) {
    if (!Number.isFinite(value) || value <= 0) return 1;
    if (value <= 1) return 1;
    const power = Math.pow(10, Math.floor(Math.log10(value)));
    const fraction = value / power;
    const nice = fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10;
    return nice * power;
  }

  function drawChartGrid(ctx, pad, w, h, max, options = {}) {
    const ticks = options.ticks || 4;
    const plotH = h - pad.top - pad.bottom;
    const plotW = w - pad.left - pad.right;
    ctx.save();
    ctx.strokeStyle = "#e8eef5";
    ctx.lineWidth = 1;
    ctx.fillStyle = "#64748b";
    ctx.font = "11px Segoe UI";
    ctx.textAlign = "right";
    ctx.textBaseline = "middle";
    for (let i = 0; i <= ticks; i += 1) {
      const ratio = i / ticks;
      const y = pad.top + plotH - ratio * plotH;
      ctx.beginPath();
      ctx.moveTo(pad.left, y);
      ctx.lineTo(w - pad.right, y);
      ctx.stroke();
      const value = max * ratio;
      const label = options.percent ? fmtPct(value) : `${fmtNumber(value, value >= 10 ? 0 : 1)}${options.suffix || ""}`;
      ctx.fillText(label, pad.left - 9, y);
    }
    ctx.strokeStyle = "#dbe5ef";
    ctx.beginPath();
    ctx.moveTo(pad.left, pad.top);
    ctx.lineTo(pad.left, h - pad.bottom);
    ctx.lineTo(w - pad.right, h - pad.bottom);
    ctx.stroke();
    if (options.yLabel) {
      ctx.textAlign = "left";
      ctx.textBaseline = "alphabetic";
      ctx.fillStyle = "#64748b";
      ctx.font = "800 11px Segoe UI";
      ctx.fillText(options.yLabel, pad.left, pad.top - 10);
    }
    ctx.restore();
    return { plotW, plotH };
  }

  function drawRoundedRect(ctx, x, y, w, h, r) {
    const radius = Math.min(r, Math.abs(w) / 2, Math.abs(h) / 2);
    ctx.beginPath();
    if (ctx.roundRect) {
      ctx.roundRect(x, y, w, h, radius);
    } else {
      ctx.moveTo(x + radius, y);
      ctx.lineTo(x + w - radius, y);
      ctx.quadraticCurveTo(x + w, y, x + w, y + radius);
      ctx.lineTo(x + w, y + h - radius);
      ctx.quadraticCurveTo(x + w, y + h, x + w - radius, y + h);
      ctx.lineTo(x + radius, y + h);
      ctx.quadraticCurveTo(x, y + h, x, y + h - radius);
      ctx.lineTo(x, y + radius);
      ctx.quadraticCurveTo(x, y, x + radius, y);
    }
  }

  function drawCanvasLegend(ctx, items, x, y) {
    ctx.save();
    ctx.font = "800 11px Segoe UI";
    ctx.textBaseline = "middle";
    items.forEach((item, index) => {
      const left = x + index * 96;
      drawRoundedRect(ctx, left, y - 5, 20, 6, 3);
      ctx.fillStyle = item.color;
      ctx.fill();
      ctx.fillStyle = "#334155";
      ctx.textAlign = "left";
      ctx.fillText(item.label, left + 28, y);
    });
    ctx.restore();
  }

  function drawGauge(id, value, label) {
    const canvas = document.getElementById(id);
    if (!canvas) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const cx = w / 2;
    const cy = h * 0.66;
    const radius = Math.min(w, h) * 0.35;
    const start = Math.PI;
    const end = Math.PI * 2;
    const clamped = Math.max(0, Math.min(Number(value) || 0, 1));
    const color = clamped >= 0.85 ? "#20b970" : clamped >= 0.75 ? "#f59e0b" : "#dc2626";

    animateChart((progress) => {
      const eased = easeOutCubic(progress);
      ctx.clearRect(0, 0, w, h);
      ctx.save();
      ctx.lineWidth = Math.max(16, radius * 0.18);
      ctx.lineCap = "round";
      ctx.beginPath();
      ctx.arc(cx, cy, radius, start, end);
      ctx.strokeStyle = "#edf3f9";
      ctx.stroke();
      ctx.beginPath();
      ctx.arc(cx, cy, radius, start, start + Math.PI * clamped * eased);
      ctx.strokeStyle = color;
      ctx.shadowColor = `${color}33`;
      ctx.shadowBlur = 10;
      ctx.stroke();
      ctx.shadowBlur = 0;
      ctx.fillStyle = "#10284b";
      ctx.font = "900 32px Segoe UI";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(fmtPct(clamped * eased), cx, cy - 14);
      ctx.fillStyle = "#64748b";
      ctx.font = "800 12px Segoe UI";
      ctx.fillText(label, cx, cy + 19);
      ctx.fillStyle = "#94a3b8";
      ctx.font = "11px Segoe UI";
      ctx.fillText("Objectif 85 %", cx, Math.min(h - 18, cy + 42));
      ctx.restore();
    }, 620);
    canvas.__chartHitAreas = [{ type: "point", x: cx, y: cy, radius: radius + 24, title: label, value: fmtPct(clamped), detail: "Indicateur recalculé automatiquement" }];
    bindChartTooltip(canvas);
  }

  function drawDonut(id, data, total) {
    const canvas = document.getElementById(id);
    if (!canvas || !data.length) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const cx = w / 2;
    const cy = h / 2;
    const radius = Math.min(w, h) * 0.34;
    const ringWidth = Math.max(28, radius * 0.34);
    const gap = 0.018;
    const segments = [];

    const renderFrame = (progress) => {
      let start = -Math.PI / 2;
      segments.length = 0;
      ctx.clearRect(0, 0, w, h);
      ctx.beginPath();
      ctx.arc(cx, cy, radius, 0, Math.PI * 2);
      ctx.strokeStyle = "#eef3f8";
      ctx.lineWidth = ringWidth;
      ctx.stroke();

      data.forEach((item, index) => {
        const rawAngle = (item.value / Math.max(total, 1)) * Math.PI * 2;
        const angle = rawAngle * progress;
        const end = start + Math.max(0, angle - gap);
        ctx.beginPath();
        ctx.arc(cx, cy, radius, start, end);
        ctx.strokeStyle = chartColor(index);
        ctx.lineWidth = ringWidth;
        ctx.lineCap = "butt";
        ctx.stroke();
        segments.push({ ...item, color: chartColor(index), start, end: start + rawAngle });
        start += rawAngle;
      });

      ctx.fillStyle = "#10284b";
      ctx.font = "800 16px Segoe UI";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText("Top 5", cx, cy - 8);
      ctx.fillStyle = "#64748b";
      ctx.font = "12px Segoe UI";
      ctx.fillText("arrêts", cx, cy + 12);
    };

    animateChart((progress) => renderFrame(easeOutCubic(progress)), 650);
    canvas.__chartSegments = segments;
    canvas.__chartCenter = { cx, cy, radius, innerRadius: radius - ringWidth / 2, outerRadius: radius + ringWidth / 2, total };
    bindDonutTooltip(canvas);
  }

  function drawDailyTrend(id, rows) {
    const canvas = document.getElementById(id);
    if (!canvas || !rows.length) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const pad = { left: 56, right: 26, top: 34, bottom: 54 };
    const plotW = w - pad.left - pad.right;
    const plotH = h - pad.top - pad.bottom;
    const max = niceMax(Math.max(...rows.map((row) => row.stopHours), 1));
    const labelStep = Math.max(1, Math.ceil(rows.length / Math.max(4, Math.floor(plotW / 86))));
    const points = rows.map((row, index) => ({
      row,
      x: pad.left + (plotW * index) / Math.max(rows.length - 1, 1),
      y: pad.top + plotH - (row.stopHours / max) * plotH
    }));

    animateChart((progress) => {
      const eased = easeOutCubic(progress);
      ctx.clearRect(0, 0, w, h);
      drawChartGrid(ctx, pad, w, h, max, { suffix: "h", yLabel: "Heures d'arrêt" });
      ctx.save();
      ctx.beginPath();
      points.forEach((point, index) => {
        const x = points[0].x + (point.x - points[0].x) * eased;
        const y = pad.top + plotH - ((point.row.stopHours * eased) / max) * plotH;
        if (index === 0) ctx.moveTo(x, pad.top + plotH);
        ctx.lineTo(x, y);
      });
      [...points].reverse().forEach((point) => {
        const x = points[0].x + (point.x - points[0].x) * eased;
        ctx.lineTo(x, pad.top + plotH);
      });
      ctx.closePath();
      const gradient = ctx.createLinearGradient(0, pad.top, 0, h - pad.bottom);
      gradient.addColorStop(0, "rgba(220, 38, 38, 0.18)");
      gradient.addColorStop(1, "rgba(220, 38, 38, 0.02)");
      ctx.fillStyle = gradient;
      ctx.fill();

      ctx.beginPath();
      points.forEach((point, index) => {
        const x = points[0].x + (point.x - points[0].x) * eased;
        const y = pad.top + plotH - ((point.row.stopHours * eased) / max) * plotH;
        if (index === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.strokeStyle = "#dc2626";
      ctx.lineWidth = 3;
      ctx.lineJoin = "round";
      ctx.stroke();

      points.forEach((point, index) => {
        const x = points[0].x + (point.x - points[0].x) * eased;
        const y = pad.top + plotH - ((point.row.stopHours * eased) / max) * plotH;
        ctx.beginPath();
        ctx.arc(x, y, 4.3, 0, Math.PI * 2);
        ctx.fillStyle = "#ffffff";
        ctx.fill();
        ctx.strokeStyle = "#dc2626";
        ctx.lineWidth = 2;
        ctx.stroke();
        if (index % labelStep === 0 || index === rows.length - 1) {
          ctx.fillStyle = "#475569";
          ctx.font = "11px Segoe UI";
          ctx.textAlign = "center";
          ctx.textBaseline = "top";
          ctx.fillText(dayLabel(`${point.row.day}T00:00:00`), point.x, h - pad.bottom + 18);
        }
      });
      ctx.restore();
    }, 640);

    canvas.__chartHitAreas = points.map((point) => ({
      type: "point",
      x: point.x,
      y: point.y,
      radius: 12,
      title: dayLabel(`${point.row.day}T00:00:00`),
      value: fmtHours(point.row.stopHours),
      detail: "Temps d'arrêt journalier"
    }));
    bindChartTooltip(canvas);
  }

  function drawCircuitBars(id, rows) {
    const canvas = document.getElementById(id);
    if (!canvas || !rows.length) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const pad = { left: 60, right: 30, top: 52, bottom: 62 };
    const plotW = w - pad.left - pad.right;
    const plotH = h - pad.top - pad.bottom;
    const slot = plotW / rows.length;
    const labelStep = slot < 74 ? 2 : 1;
    const hitAreas = [];

    animateChart((progress) => {
      hitAreas.length = 0;
      ctx.clearRect(0, 0, w, h);
      drawChartGrid(ctx, pad, w, h, 1, { percent: true, yLabel: "TRS (%)" });
      rows.forEach((row, index) => {
        const barW = Math.min(58, Math.max(32, slot * 0.46));
        const x = pad.left + index * slot + (slot - barW) / 2;
        const barH = Math.max(2, row.value * plotH * easeOutCubic(progress));
        const y = pad.top + plotH - barH;
        const color = chartColor(index + 4);
        const gradient = ctx.createLinearGradient(0, y, 0, pad.top + plotH);
        gradient.addColorStop(0, color);
        gradient.addColorStop(1, `${color}bb`);
        drawRoundedRect(ctx, x, y, barW, barH, 7);
        ctx.fillStyle = gradient;
        ctx.fill();
        ctx.fillStyle = "#10284b";
        ctx.font = "800 12px Segoe UI";
        ctx.textAlign = "center";
        ctx.textBaseline = "alphabetic";
        const labelY = Math.max(pad.top - 4, y - 8);
        ctx.fillText(fmtPct(row.value), x + barW / 2, labelY);
        if (index % labelStep === 0) {
          ctx.fillStyle = "#475569";
          ctx.font = "12px Segoe UI";
          ctx.fillText(truncate(row.label, labelStep > 1 ? 10 : 14), x + barW / 2, h - 22);
        }
        hitAreas.push({ type: "rect", x, y, w: barW, h: barH, title: row.label, value: fmtPct(row.value), detail: "Performance du circuit" });
      });
    }, 520);
    canvas.__chartHitAreas = hitAreas;
    bindChartTooltip(canvas);
  }

  function drawPareto(id, data, total) {
    const canvas = document.getElementById(id);
    if (!canvas || !data.length) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const pad = { left: 62, right: 64, top: 42, bottom: 70 };
    const plotW = w - pad.left - pad.right;
    const plotH = h - pad.top - pad.bottom;
    const max = niceMax(Math.max(...data.map((d) => d.value), 1));
    const slot = plotW / data.length;
    const labelStep = slot < 58 ? Math.ceil(58 / slot) : 1;
    const hitAreas = [];

    animateChart((progress) => {
      let cumulative = 0;
      const points = [];
      hitAreas.length = 0;
      ctx.clearRect(0, 0, w, h);
      drawChartGrid(ctx, pad, w, h, max, { suffix: "h", yLabel: "Durée cumulée" });
      data.forEach((item, i) => {
        const barW = Math.max(14, Math.min(42, slot * 0.54));
        const x = pad.left + i * slot + (slot - barW) / 2;
        const barH = (item.value / max) * plotH * easeOutCubic(progress);
        const y = pad.top + plotH - barH;
        const color = i < 5 ? chartColor(i) : "#94a3b8";
        const gradient = ctx.createLinearGradient(0, y, 0, pad.top + plotH);
        gradient.addColorStop(0, color);
        gradient.addColorStop(1, `${color}aa`);
        drawRoundedRect(ctx, x, y, barW, barH, 6);
        ctx.fillStyle = gradient;
        ctx.fill();
        cumulative += item.value;
        const share = total ? cumulative / total : 0;
        points.push({
          x: x + barW / 2,
          y: pad.top + plotH - share * progress * plotH,
          share
        });
        if (i % labelStep === 0) {
          ctx.fillStyle = "#475569";
          ctx.font = "11px Segoe UI";
          ctx.textAlign = "center";
          ctx.textBaseline = "top";
          ctx.fillText(truncate(item.label, labelStep > 1 ? 8 : 12), x + barW / 2, h - pad.bottom + 18);
        }
        hitAreas.push({
          type: "rect",
          x,
          y,
          w: barW,
          h: barH,
          title: item.label,
          value: fmtHours(item.value),
          detail: `Part: ${fmtPct(item.value / Math.max(total, 1))}`
        });
      });
      ctx.beginPath();
      points.forEach((point, index) => {
        if (index === 0) ctx.moveTo(point.x, point.y);
        else ctx.lineTo(point.x, point.y);
      });
      ctx.strokeStyle = "#f97316";
      ctx.lineWidth = 2.5;
      ctx.stroke();
      points.forEach((point) => {
        ctx.beginPath();
        ctx.arc(point.x, point.y, 3.5, 0, Math.PI * 2);
        ctx.fillStyle = "#ffffff";
        ctx.fill();
        ctx.strokeStyle = "#f97316";
        ctx.lineWidth = 2;
        ctx.stroke();
      });
      ctx.fillStyle = "#64748b";
      ctx.font = "11px Segoe UI";
      ctx.textAlign = "right";
      ctx.fillText("% cumulé", w - pad.right, pad.top - 14);
      drawCanvasLegend(ctx, [
        { label: "Heures", color: "#2563eb" },
        { label: "% cumulé", color: "#f97316" }
      ], Math.max(pad.left, w - pad.right - 184), pad.top - 18);
    }, 560);
    canvas.__chartHitAreas = hitAreas;
    bindChartTooltip(canvas);
  }

  function drawBars(id, data, options = {}) {
    const canvas = document.getElementById(id);
    if (!canvas || !data.length) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const pad = { left: 58, right: 28, top: 40, bottom: 66 };
    const plotW = w - pad.left - pad.right;
    const plotH = h - pad.top - pad.bottom;
    const max = niceMax(Math.max(...data.map((d) => d.value), 1));
    const slot = plotW / data.length;
    const labelStep = slot < 54 ? Math.ceil(54 / slot) : 1;
    const hitAreas = [];
    animateChart((progress) => {
      hitAreas.length = 0;
      ctx.clearRect(0, 0, w, h);
      drawChartGrid(ctx, pad, w, h, max, { suffix: options.suffix || "", yLabel: options.yLabel || "Valeur" });
      data.forEach((item, i) => {
        const barW = Math.max(9, Math.min(42, slot * 0.58));
        const x = pad.left + i * slot + (slot - barW) / 2;
        const barH = (item.value / max) * plotH * easeOutCubic(progress);
        const y = pad.top + plotH - barH;
        const color = options.color || "#0f766e";
        const gradient = ctx.createLinearGradient(0, y, 0, pad.top + plotH);
        gradient.addColorStop(0, color);
        gradient.addColorStop(1, `${color}aa`);
        drawRoundedRect(ctx, x, y, barW, barH, 6);
        ctx.fillStyle = gradient;
        ctx.fill();
        if (slot > 48 && item.value > 0) {
          ctx.fillStyle = "#10284b";
          ctx.font = "800 11px Segoe UI";
          ctx.textAlign = "center";
          ctx.textBaseline = "bottom";
          ctx.fillText(`${fmtNumber(item.value, item.value >= 10 ? 0 : 1)}${options.suffix || ""}`, x + barW / 2, Math.max(14, y - 6));
        }
        if (i % labelStep === 0) {
          ctx.fillStyle = "#475569";
          ctx.font = "11px Segoe UI";
          ctx.textAlign = "center";
          ctx.textBaseline = "top";
          ctx.fillText(truncate(item.label, labelStep > 1 ? 8 : 11), x + barW / 2, h - pad.bottom + 18);
        }
        hitAreas.push({
          type: "rect",
          x,
          y,
          w: barW,
          h: barH,
          title: item.label,
          value: `${fmtNumber(item.value, item.value >= 10 ? 0 : 1)}${options.suffix || ""}`,
          detail: options.detail || "Indicateur opérationnel"
        });
      });
    }, 520);
    canvas.__chartHitAreas = hitAreas;
    bindChartTooltip(canvas);
  }

  function drawLineBars(id, rows) {
    const canvas = document.getElementById(id);
    if (!canvas || !rows.length) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const pad = { left: 66, right: 34, top: 46, bottom: 72 };
    const plotW = w - pad.left - pad.right;
    const plotH = h - pad.top - pad.bottom;
    const max = niceMax(Math.max(...rows.flatMap((r) => [r.bar, r.line]), 1));
    const slot = plotW / rows.length;
    const barW = Math.max(7, Math.min(34, slot * 0.5));
    const labelStep = slot < 56 ? Math.ceil(56 / slot) : 1;
    const hitAreas = [];

    animateChart((progress) => {
      const eased = easeOutCubic(progress);
      const linePoints = [];
      hitAreas.length = 0;
      ctx.clearRect(0, 0, w, h);
      drawChartGrid(ctx, pad, w, h, max, { suffix: "t", yLabel: "Tonnage" });
      drawCanvasLegend(ctx, [
        { label: "Pesage", color: "#0f766e" },
        { label: "Draft", color: "#f97316" }
      ], Math.max(pad.left, w - pad.right - 190), pad.top - 20);

      rows.forEach((row, i) => {
        const x = pad.left + i * slot + (slot - barW) / 2;
        const barH = (row.bar / max) * plotH * eased;
        const y = pad.top + plotH - barH;
        const gradient = ctx.createLinearGradient(0, y, 0, pad.top + plotH);
        gradient.addColorStop(0, "#0f766e");
        gradient.addColorStop(1, "rgba(15, 118, 110, 0.62)");
        drawRoundedRect(ctx, x, y, barW, barH, 6);
        ctx.fillStyle = gradient;
        ctx.fill();
        const point = {
          x: x + barW / 2,
          y: pad.top + plotH - ((row.line * eased) / max) * plotH,
          row
        };
        linePoints.push(point);
        if (i % labelStep === 0) {
          ctx.fillStyle = "#475569";
          ctx.font = "11px Segoe UI";
          ctx.textAlign = "center";
          ctx.textBaseline = "top";
          ctx.fillText(truncate(row.label, 10), x + barW / 2, h - pad.bottom + 18);
        }
        hitAreas.push({
          type: "rect",
          x,
          y,
          w: barW,
          h: barH,
          title: row.label,
          value: `Pesage: ${fmtNumber(row.bar, 0)} t`,
          detail: `Draft: ${fmtNumber(row.line, 0)} t`
        });
      });

      ctx.beginPath();
      linePoints.forEach((point, i) => {
        if (i === 0) ctx.moveTo(point.x, point.y);
        else ctx.lineTo(point.x, point.y);
      });
      ctx.strokeStyle = "#f97316";
      ctx.lineWidth = 3;
      ctx.lineJoin = "round";
      ctx.stroke();
      linePoints.forEach((point) => {
        ctx.beginPath();
        ctx.arc(point.x, point.y, 4, 0, Math.PI * 2);
        ctx.fillStyle = "#ffffff";
        ctx.fill();
        ctx.strokeStyle = "#f97316";
        ctx.lineWidth = 2;
        ctx.stroke();
        hitAreas.push({
          type: "point",
          x: point.x,
          y: point.y,
          radius: 10,
          title: point.row.label,
          value: `Draft: ${fmtNumber(point.row.line, 0)} t`,
          detail: `Pesage: ${fmtNumber(point.row.bar, 0)} t`
        });
      });
    }, 640);
    canvas.__chartHitAreas = hitAreas;
    bindChartTooltip(canvas);
  }

  function setupCanvas(canvas) {
    const ratioValue = pixelRatio();
    const rect = canvas.getBoundingClientRect();
    canvas.width = Math.max(320, rect.width) * ratioValue;
    canvas.height = Math.max(220, rect.height || canvas.clientHeight) * ratioValue;
    const ctx = canvas.getContext("2d");
    ctx.setTransform(ratioValue, 0, 0, ratioValue, 0, 0);
    return ctx;
  }

  function drawAxis(ctx, pad, w, h) {
    ctx.strokeStyle = "#d7ded6";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(pad.left, pad.top);
    ctx.lineTo(pad.left, h - pad.bottom);
    ctx.lineTo(w - pad.right, h - pad.bottom);
    ctx.stroke();
  }

  function exportEventsCsv() {
    const headers = ["id", "sectionKey", "subEquipment", "family", "start", "end", "durationHours", "description", "assignment", "quality", "destination"];
    const rows = getFilteredEvents().map((event) => headers.map((key) => csvCell(event[key])).join(";"));
    downloadFile("trace-port-arrets.csv", [headers.join(";"), ...rows].join("\n"), "text/csv;charset=utf-8");
  }

  function exportSummaryJson() {
    const events = getAnalysisEvents();
    const summary = {
      generatedAt: new Date().toISOString(),
      sourceWorkbook: DATA.sourceWorkbook,
      filters: state.filters,
      metrics: computeMetrics(events),
      trains: getAllTrains(),
      ships: getAllShips(),
      dailySynthesis: getAllDays().map(computeDaySummary),
      topFamilies: topGroups(groupSum(events, "family"), 20),
      topSections: topGroups(groupSum(events, "sectionKey", "Non affecté"), 20)
    };
    downloadFile("trace-port-synthese.json", JSON.stringify(summary, null, 2), "application/json;charset=utf-8");
  }

  function downloadFile(filename, content, type) {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  }

  function groupSum(rows, key, blankLabel = "-") {
    const grouped = new Map();
    rows.forEach((row) => {
      const label = row[key] || blankLabel;
      grouped.set(label, (grouped.get(label) || 0) + (Number(row.durationHours) || 0));
    });
    return grouped;
  }

  function topGroups(grouped, limit = 10) {
    return Array.from(grouped.entries())
      .map(([label, value]) => ({ label, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, limit);
  }

  function sumByFamily(events, families) {
    const wanted = new Set(families.map(normalize));
    return events.reduce((acc, event) => wanted.has(normalize(event.family)) ? acc + (Number(event.durationHours) || 0) : acc, 0);
  }

  function sum(rows, key) {
    return rows.reduce((acc, row) => acc + (Number(row[key]) || 0), 0);
  }

  function average(values) {
    if (!values.length) return 0;
    return values.reduce((acc, value) => acc + value, 0) / values.length;
  }

  function ratio(numerator, denominator) {
    if (!denominator) return 0;
    return Math.max(0, numerator / denominator);
  }

  function hoursBetweenLocal(start, end) {
    if (!start || !end) return 0;
    const ms = new Date(end).getTime() - new Date(start).getTime();
    return Math.round((ms / 36e5) * 100) / 100;
  }

  function numberFromInput(value) {
    const number = Number(value);
    return Number.isFinite(number) ? number : 0;
  }

  function dateKey(value) {
    if (!value) return "";
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    return String(value).slice(0, 10);
  }

  function getAllDays() {
    return unique([
      ...getAllEvents().map((event) => dateKey(event.start)),
      ...getAllTrains().map((train) => dateKey(train.day)),
      ...getAllShips().map((ship) => dateKey(ship.start)),
      ...DATA.tonnage.map((row) => dateKey(row.day))
    ]).filter(Boolean);
  }

  function unique(values) {
    return Array.from(new Set(values.map((value) => String(value).trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, "fr"));
  }

  function hashString(value) {
    return String(value || "").split("").reduce((acc, char) => {
      return ((acc << 5) - acc + char.charCodeAt(0)) >>> 0;
    }, 0);
  }

  function normalize(value) {
    return String(value || "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  }

  function fmtCompactEur(value) {
    const n = Number(value);
    if (!Number.isFinite(n) || n === 0) return "0 €";
    const abs = Math.abs(n);
    if (abs >= 1e9) return `${(n / 1e9).toLocaleString("fr-FR", { maximumFractionDigits: 2 })} Md €`;
    if (abs >= 1e6) return `${(n / 1e6).toLocaleString("fr-FR", { maximumFractionDigits: 2 })} M €`;
    if (abs >= 1e3) return `${(n / 1e3).toLocaleString("fr-FR", { maximumFractionDigits: 1 })} k €`;
    return `${n.toLocaleString("fr-FR", { maximumFractionDigits: 0 })} €`;
  }

  function fmtNumber(value, digits = 1) {
    const number = Number(value);
    if (!Number.isFinite(number)) return "-";
    return number.toLocaleString("fr-FR", { maximumFractionDigits: digits, minimumFractionDigits: 0 });
  }

  function fmtHours(value) {
    const number = Number(value);
    if (!Number.isFinite(number) || number === 0) return "0 h";
    return `${number.toLocaleString("fr-FR", { maximumFractionDigits: 1 })} h`;
  }

  function fmtHoursOrNumber(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) return "-";
    return number > 1000 ? fmtNumber(number, 0) : fmtHours(number);
  }

  function fmtPct(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) return "-";
    return `${(number * 100).toLocaleString("fr-FR", { maximumFractionDigits: 1 })} %`;
  }

  function fmtDate(value) {
    if (!value) return "-";
    return new Date(value).toLocaleDateString("fr-FR", { day: "2-digit", month: "2-digit", year: "numeric" });
  }

  function fmtDateFromKey(value) {
    if (!value) return "-";
    return new Date(`${value}T00:00:00`).toLocaleDateString("fr-FR", { day: "2-digit", month: "2-digit", year: "numeric" });
  }

  function fmtDateTime(value) {
    if (!value) return "-";
    return new Date(value).toLocaleString("fr-FR", {
      day: "2-digit",
      month: "2-digit",
      hour: "2-digit",
      minute: "2-digit"
    });
  }

  function dayLabel(value) {
    if (!value) return "-";
    return new Date(value).toLocaleDateString("fr-FR", { day: "2-digit" });
  }

  function truncate(value, length) {
    const text = String(value || "");
    return text.length > length ? `${text.slice(0, length - 1)}…` : text;
  }

  function csvCell(value) {
    const text = String(value ?? "").replace(/"/g, '""');
    return `"${text}"`;
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function escapeAttr(value) {
    return escapeHtml(value);
  }

  function pixelRatio() {
    return window.devicePixelRatio || 1;
  }

  init();
})();
