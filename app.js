(function () {
  "use strict";

  const DATA = window.PFE_DATA;
  const EVENTS_KEY = "trace-port-digital-events-v1";
  const TRAINS_KEY = "trace-port-digital-trains-v1";
  const SHIPS_KEY = "trace-port-digital-ships-v1";
  const VALIDATIONS_KEY = "trace-port-validations-v1";
  const LOGS_KEY = "trace-port-logs-v1";
  const CURRENT_USER = {
    name: "El Houdzi Israa",
    role: "Administrateur",
    team: "Direction Logistique Portuaire"
  };
  const VIEW_TITLES = {
    dashboard: "Tableau de bord",
    entry: "Nouvel arrêt",
    myStops: "Mes arrêts",
    currentStops: "Arrêts en cours",
    validation: "Validation",
    kpiDashboard: "Tableaux de bord",
    indicators: "Indicateurs (KPI)",
    pareto: "Pareto des arrêts",
    performance: "Performance circuits",
    dailyReports: "Rapports journaliers",
    monthlyReports: "Rapports mensuels",
    exportData: "Export des données",
    equipments: "Équipements",
    stopNatures: "Natures d'arrêt",
    users: "Utilisateurs",
    settings: "Paramètres",
    logs: "Traçabilité (logs)",
    daily: "Synthèse journalière",
    monthly: "Synthèse mensuelle",
    events: "Bilan des arrêts",
    tonnage: "Tonnage",
    flow: "Trains & navires",
    formulas: "Formules & requêtes",
    dmaic: "Besoins PFE"
  };
  const VIEW_META = {
    dashboard: ["Monitoring", "Vue consolidée temps réel", "Surveiller les alertes et ouvrir les arrêts critiques", "currentStops"],
    entry: ["Opérationnel", "Déclaration d'un arrêt", "Enregistrer l'arrêt pour alimenter Mes arrêts et Validation", "myStops"],
    myStops: ["Opérationnel", "Suivi des arrêts déclarés", "Ouvrir le détail ou suivre le statut de validation", "validation"],
    currentStops: ["Opérationnel", "Incidents actifs et critiques", "Prioriser les arrêts longs et déclencher validation", "validation"],
    validation: ["Opérationnel", "Contrôle chef d'équipe", "Valider ou rejeter pour recalculer les KPI", "kpiDashboard"],
    stopDetail: ["Opérationnel", "Fiche complète d'incident", "Consulter l'historique puis valider ou corriger", "validation"],
    kpiDashboard: ["Monitoring", "Pilotage de performance", "Analyser les écarts puis ouvrir les indicateurs", "indicators"],
    indicators: ["Monitoring", "KPI industriels", "Comparer TRS, TRG, MTTR, MTBF et débit", "performance"],
    performance: ["Monitoring", "Performance par circuit", "Identifier les circuits sous objectif", "pareto"],
    pareto: ["Décisionnel", "Analyse des causes", "Prioriser les causes dominantes", "dailyReports"],
    dailyReports: ["Décisionnel", "Rapport journalier", "Générer et diffuser la synthèse quotidienne", "monthlyReports"],
    monthlyReports: ["Décisionnel", "Rapport mensuel", "Consolider la fiche mensuelle envoyée aux services", "exportData"],
    exportData: ["Décisionnel", "Extraction contrôlée", "Exporter les données validées", "logs"],
    equipments: ["Administration", "Référentiel équipements", "Maintenir les circuits et équipements", "stopNatures"],
    stopNatures: ["Administration", "Référentiel natures d'arrêt", "Normaliser les familles de causes", "users"],
    users: ["Administration", "Utilisateurs et rôles", "Gérer les accès RBAC", "logs"],
    logs: ["Administration", "Audit et traçabilité", "Contrôler les actions sensibles", "settings"],
    settings: ["Administration", "Configuration plateforme", "Piloter les paramètres système", "dashboard"]
  };
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

  const state = {
    view: "dashboard",
    filters: {
      section: "all",
      family: "all",
      quality: "all",
      search: ""
    },
    dailyDate: null,
    selectedEventId: null,
    formulaSearch: "",
    formulaSheet: "all"
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
    dataCount: document.getElementById("data-count")
  };

  function init() {
    els.dataCount.textContent = `${DATA.events.length} lignes Bilan, ${DATA.formulasCount} formules`;
    populateFilters();
    bindShell();
    render();
  }

  function bindShell() {
    document.querySelectorAll(".nav-item").forEach((button) => {
      button.addEventListener("click", () => {
        state.view = button.dataset.view;
        document.querySelectorAll(".nav-item").forEach((b) => b.classList.toggle("active", b === button));
        render();
      });
    });

    [els.section, els.family, els.quality].forEach((select) => {
      select.addEventListener("change", () => {
        state.filters.section = els.section.value;
        state.filters.family = els.family.value;
        state.filters.quality = els.quality.value;
        render();
      });
    });

    els.search.addEventListener("input", () => {
      state.filters.search = els.search.value.trim().toLowerCase();
      if (els.globalSearch && els.globalSearch.value !== els.search.value) {
        els.globalSearch.value = els.search.value;
      }
      render();
    });

    els.globalSearch?.addEventListener("input", () => {
      state.filters.search = els.globalSearch.value.trim().toLowerCase();
      els.search.value = els.globalSearch.value;
      render();
    });

    els.globalDate?.addEventListener("change", () => {
      state.dailyDate = els.globalDate.value || state.dailyDate;
      if (state.view === "dailyReports") render();
    });

    document.getElementById("export-csv").addEventListener("click", exportEventsCsv);
    document.getElementById("export-json").addEventListener("click", exportSummaryJson);
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

  function addLog(action, objectId, detail, user = CURRENT_USER.name) {
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
      return sectionOk && familyOk && qualityOk && (!query || haystack.includes(query));
    });
  }

  function getAnalysisEvents() {
    return getFilteredEvents().filter((event) => {
      const decorated = decorateEvent(event);
      if (String(event.id || "").startsWith("LOCAL-")) return decorated.status === "validated";
      return decorated.status !== "rejected";
    });
  }

  function render() {
    const title = VIEW_TITLES[state.view] || "Tableau de bord";
    els.title.textContent = title;

    const renderers = {
      dashboard: renderDashboard,
      entry: renderEntry,
      myStops: renderMyStops,
      currentStops: renderCurrentStopsView,
      validation: renderValidation,
      stopDetail: renderStopDetail,
      kpiDashboard: renderKpiDashboard,
      indicators: renderIndicators,
      pareto: renderParetoAnalysis,
      performance: renderPerformanceCircuits,
      dailyReports: renderReports,
      monthlyReports: renderReports,
      exportData: renderExportData,
      equipments: renderEquipments,
      stopNatures: renderStopNatures,
      users: renderUsers,
      settings: renderSettings,
      logs: renderLogs,
      daily: renderDailySynthesis,
      monthly: renderMonthlySynthesis,
      events: renderEvents,
      tonnage: renderTonnage,
      flow: renderFlow,
      formulas: renderFormulas,
      dmaic: renderDmaic
    };
    (renderers[state.view] || renderDashboard)();
    injectProcessContext();
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
        <div class="process-flow" aria-label="Flux opérationnel">
          ${workflowStep("Saisie", ["entry", "myStops"].includes(state.view))}
          ${workflowStep("Validation", ["validation", "stopDetail"].includes(state.view))}
          ${workflowStep("KPI", ["dashboard", "kpiDashboard", "indicators", "performance"].includes(state.view))}
          ${workflowStep("Décision", ["pareto", "dailyReports", "monthlyReports", "exportData"].includes(state.view))}
          ${workflowStep("Audit", ["equipments", "stopNatures", "users", "logs", "settings"].includes(state.view))}
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

  function workflowStep(label, active) {
    return `<span class="workflow-step ${active ? "active" : ""}">${escapeHtml(label)}</span>`;
  }

  function renderDashboard() {
    const sourceEvents = getFilteredEvents();
    const events = getAnalysisEvents();
    const metrics = computeMetrics(events);
    const pareto = topGroups(groupSum(events, "family"), 5);
    const trend = getAllDays().map(computeDaySummary).slice(-7);
    const alerts = buildDashboardAlerts(metrics, pareto, sourceEvents);
    const currentStops = decorateEvents(sourceEvents).filter((event) => event.status === "pending" || Number(event.durationHours) >= 2).slice(0, 3);
    const circuitRows = buildCircuitPerformance(metrics);

    els.view.innerHTML = `
      <div class="metric-grid dashboard-kpis">
        ${kpiCard("TRS Global", fmtPct(metrics.trsGlobal), "Objectif : 85%", "green")}
        ${kpiCard("TRS Exploitation", fmtPct(metrics.trsExploitation), "Objectif : 85%", "blue")}
        ${kpiCard("TRS Maintenance", fmtPct(metrics.trsMaintenance), "Objectif : 80%", "purple")}
        ${kpiCard("Temps d'arrêt total", fmtHours(metrics.totalStopHours), "Cumul mensuel", "amber")}
        ${kpiCard("Nombre d'arrêts", fmtNumber(events.length, 0), "Lignes intégrées", "red")}
        ${kpiCard("Débit global", `${fmtNumber(metrics.cadenceTph, 0)} t/h`, "Tonnage / marche", "teal")}
      </div>

      <div class="dashboard-main">
        <section class="panel">
          <div class="panel-head">
            <h2>Arrêts par nature (Top 5)</h2>
            <span class="badge">${fmtHours(metrics.totalStopHours)}</span>
          </div>
          <div class="donut-layout">
            <canvas id="nature-donut" class="mini-chart donut-chart"></canvas>
            ${renderDonutLegend(pareto, metrics.totalStopHours)}
          </div>
        </section>

        <section class="panel">
          <div class="panel-head">
            <h2>Evolution du temps d'arrêt</h2>
            <span class="badge blue">7 derniers jours</span>
          </div>
          <canvas id="daily-trend-chart" class="chart"></canvas>
        </section>

        <section class="panel">
          <div class="panel-head">
            <h2>Alertes</h2>
            <span class="badge red">${alerts.length}</span>
          </div>
          ${renderAlerts(alerts)}
        </section>
      </div>

      <div class="dashboard-bottom">
        <section class="panel">
          <div class="panel-head">
            <h2>Arrêts critiques</h2>
            <span class="badge warn">${currentStops.length}</span>
          </div>
          ${renderCurrentStops(currentStops)}
        </section>

        <section class="panel">
          <div class="panel-head">
            <h2>Performance par circuit</h2>
            <span class="badge cyan">TRS</span>
          </div>
          <canvas id="circuit-chart" class="mini-chart"></canvas>
        </section>

        <section class="panel">
          <div class="panel-head">
            <h2>Actions rapides</h2>
          </div>
          ${renderQuickActions()}
        </section>
      </div>
    `;

    bindQuickActions();
    requestAnimationFrame(() => {
      drawDonut("nature-donut", pareto, metrics.totalStopHours);
      drawDailyTrend("daily-trend-chart", trend);
      drawCircuitBars("circuit-chart", circuitRows);
    });
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
    const events = getAnalysisEvents();
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
          ${renderEventsTable(events.slice(0, 350))}
        </section>
        <section class="panel">
          <h2>Répartition</h2>
          <canvas id="events-family-chart" class="mini-chart"></canvas>
          <div class="split-list">${renderProgressList(bySection, sum(events, "durationHours"))}</div>
        </section>
      </div>
    `;

    requestAnimationFrame(() => {
      drawBars("events-family-chart", byFamily, { color: "#0f766e", suffix: "h" });
    });
  }

  function renderMyStops() {
    const events = decorateEvents(getFilteredEvents());
    const myEvents = events
      .filter((event) => event.declaredBy === CURRENT_USER.name || String(event.id).startsWith("LOCAL-"))
      .slice(0, 180);
    const rows = myEvents.length ? myEvents : events.slice(0, 80);
    const pending = rows.filter((event) => event.status === "pending").length;
    const validated = rows.filter((event) => event.status === "validated").length;
    const rejected = rows.filter((event) => event.status === "rejected").length;

    els.view.innerHTML = `
      <div class="metric-grid">
        ${metric("Mes arrêts", fmtNumber(rows.length, 0), "Arrêts saisis ou affectés")}
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
        ${renderOperationalStopsTable(rows, { actions: true })}
      </section>
    `;

    bindQuickActions();
    bindStopActions();
  }

  function renderCurrentStopsView() {
    const events = decorateEvents(getFilteredEvents())
      .filter((event) => event.status === "pending" || Number(event.durationHours) >= 1)
      .sort((a, b) => (b.durationHours || 0) - (a.durationHours || 0))
      .slice(0, 120);
    const critical = events.filter((event) => Number(event.durationHours) >= 2).length;
    const totalHours = sum(events, "durationHours");

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
        ${renderOperationalStopsTable(events, { actions: true, validationActions: true })}
      </section>
    `;

    bindStopActions();
  }

  function renderValidation() {
    const events = decorateEvents(getFilteredEvents());
    const pending = events.filter((event) => event.status === "pending").slice(0, 150);
    const validated = events.filter((event) => event.status === "validated").length;
    const rejected = events.filter((event) => event.status === "rejected").length;

    els.view.innerHTML = `
      <div class="metric-grid">
        ${metric("En attente", fmtNumber(pending.length, 0), "À traiter")}
        ${metric("Validés aujourd'hui", fmtNumber(Math.min(validated, 12), 0), "Workflow chef d'équipe")}
        ${metric("Rejetés aujourd'hui", fmtNumber(Math.min(rejected, 5), 0), "Avec commentaire")}
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
        ${renderOperationalStopsTable(pending, { actions: true, validationActions: true })}
      </section>
    `;

    bindStopActions();
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

  function renderKpiDashboard() {
    const events = getAnalysisEvents();
    const metrics = computeMetrics(events);
    const dailyRows = getAllDays().map(computeDaySummary).slice(-12);

    els.view.innerHTML = `
      <div class="metric-grid dashboard-kpis">
        ${kpiCard("TRS Global", fmtPct(metrics.trsGlobal), "Objectif : 85%", "green")}
        ${kpiCard("TRS Exploitation", fmtPct(metrics.trsExploitation), "Objectif : 85%", "blue")}
        ${kpiCard("TRS Maintenance", fmtPct(metrics.trsMaintenance), "Objectif : 80%", "purple")}
        ${kpiCard("TRG Global", fmtPct(metrics.trgGlobal), "Disponibilité réelle", "teal")}
        ${kpiCard("MTTR", fmtHours(computeMttr(events)), "Temps moyen réparation", "amber")}
        ${kpiCard("MTBF", fmtHours(computeMtbf(events)), "Temps moyen entre arrêts", "green")}
      </div>
      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Jauge TRS global</h2>
            <span class="badge">${fmtPct(metrics.trsGlobal)}</span>
          </div>
          <canvas id="trs-gauge" class="mini-chart"></canvas>
        </section>
        <section class="panel">
          <div class="panel-head">
            <h2>Tendance KPI</h2>
            <span class="badge blue">12 derniers jours</span>
          </div>
          <canvas id="kpi-trend" class="chart"></canvas>
        </section>
      </div>
    `;

    requestAnimationFrame(() => {
      drawGauge("trs-gauge", metrics.trsGlobal, "TRS global");
      drawDailyTrend("kpi-trend", dailyRows);
    });
  }

  function renderIndicators() {
    const events = getAnalysisEvents();
    const metrics = computeMetrics(events);
    const indicators = [
      ["TRS Global", fmtPct(metrics.trsGlobal), "Disponibilité nette après arrêts exploitation et maintenance"],
      ["TRS Exploitation", fmtPct(metrics.trsExploitation), "Impact des arrêts exploitation"],
      ["TRS Maintenance", fmtPct(metrics.trsMaintenance), "Impact maintenance électrique, instrumentation, mécanique et bande"],
      ["TRG Global", fmtPct(metrics.trgGlobal), "Temps de marche / temps disponible"],
      ["Débit global", `${fmtNumber(metrics.cadenceTph, 0)} t/h`, "Tonnage chargé / heures de marche"],
      ["Temps d'arrêt", fmtHours(metrics.totalStopHours), "Somme des arrêts filtrés"],
      ["MTTR", fmtHours(computeMttr(events)), "Durée moyenne des interventions"],
      ["MTBF", fmtHours(computeMtbf(events)), "Temps moyen entre deux arrêts"]
    ];

    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Indicateurs de performance</h2>
            <p class="status-line">Calcul automatique à partir des arrêts, tonnages, trains et navires intégrés depuis Excel ou saisis dans TRACE-PORT.</p>
          </div>
          <span class="badge cyan">Temps réel</span>
        </div>
        <div class="indicator-grid">
          ${indicators.map(([label, value, body]) => `
            <article class="indicator-card">
              <span>${escapeHtml(label)}</span>
              <strong>${value}</strong>
              <p>${escapeHtml(body)}</p>
            </article>
          `).join("")}
        </div>
      </section>
      <section class="panel">
        <div class="panel-head">
          <h2>Détail des calculs</h2>
          <button class="ghost-button" type="button" data-target-view="formulas">Voir les formules Excel reprises</button>
        </div>
        ${renderQualitySummary(metrics)}
      </section>
    `;

    bindQuickActions();
  }

  function renderParetoAnalysis() {
    const events = getAnalysisEvents();
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
    const events = getAnalysisEvents();
    const metrics = computeMetrics(events);
    const circuits = buildCircuitPerformance(metrics);
    const chargingRows = buildSynthesisRows(events, CHARGING_SECTIONS, metrics.chargingAvailableHours);
    const dischargeRows = buildSynthesisRows(events, DISCHARGE_SECTIONS, Math.max(getAllDays().length * 48, 1));

    els.view.innerHTML = `
      <div class="metric-grid">
        ${circuits.map((row, index) => kpiCard(row.label, fmtPct(row.value), "Performance circuit", ["green", "blue", "amber", "teal"][index] || "blue")).join("")}
      </div>
      <div class="two-col">
        <section class="panel">
          <div class="panel-head">
            <h2>Performance circuits</h2>
            <span class="badge cyan">TRS</span>
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
          <h2>Synthèse déchargement trains</h2>
          <span class="badge">DA10 / DB10</span>
        </div>
        ${renderSynthesisMini(dischargeRows, Math.max(getAllDays().length * 48, 1))}
      </section>
    `;

    requestAnimationFrame(() => drawCircuitBars("performance-circuit-chart", circuits));
  }

  function renderExportData() {
    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Export des données</h2>
            <p class="status-line">Exports opérationnels pour remplacer les extractions Excel manuelles et alimenter les rapports.</p>
          </div>
          <span class="badge">${fmtNumber(getAllEvents().length, 0)} arrêts</span>
        </div>
        <div class="export-grid">
          ${exportCard("Journal des arrêts", "CSV compatible Excel", "Exporter CSV", "events-csv")}
          ${exportCard("Synthèse complète", "JSON avec KPI, trains, navires et Pareto", "Exporter JSON", "summary-json")}
          ${exportCard("Rapport journalier", "HTML imprimable / PDF navigateur", "Générer", "report-daily")}
          ${exportCard("Rapport mensuel", "HTML imprimable / PDF navigateur", "Générer", "report-monthly")}
        </div>
      </section>
    `;

    bindExportCards();
  }

  function renderReports() {
    const isMonthly = state.view === "monthlyReports";
    const reportType = isMonthly ? "mensuel" : "journalier";
    const rows = buildReportRows(isMonthly ? "monthly" : "daily");

    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2>Rapports ${isMonthly ? "mensuels" : "journaliers"}</h2>
            <p class="status-line">Génération automatique à partir des synthèses recalculées par TRACE-PORT.</p>
          </div>
          <div class="inline-actions">
            <button class="primary-button" type="button" data-report-kind="${isMonthly ? "monthly" : "daily"}" data-report-format="html">Générer PDF</button>
            <button class="ghost-button" type="button" data-report-kind="${isMonthly ? "monthly" : "daily"}" data-report-format="csv">Générer Excel</button>
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
                    <button class="table-action" type="button" data-report-kind="${isMonthly ? "monthly" : "daily"}" data-report-format="html">PDF</button>
                    <button class="table-action" type="button" data-report-kind="${isMonthly ? "monthly" : "daily"}" data-report-format="csv">Excel</button>
                  </td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </section>
    `;

    bindReportButtons();
  }

  function renderEquipments() {
    const equipments = buildEquipmentRows();
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
              ${equipments.map((row) => `
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
      </section>
    `;
    bindReferenceActions();
  }

  function renderStopNatures() {
    const totals = groupSum(getFilteredEvents(), "family");
    const families = unique([...DATA.families.map((f) => f.name), ...Array.from(totals.keys())]);
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
              ${families.map((family, index) => {
                const source = DATA.families.find((item) => normalize(item.name) === normalize(family));
                return `
                  <tr>
                    <td>NA-${String(index + 1).padStart(2, "0")}</td>
                    <td>${escapeHtml(family)}</td>
                    <td>${escapeHtml(stopCategory(family))}</td>
                    <td>${fmtHours(totals.get(family) || 0)}</td>
                    <td>${escapeHtml(source?.examples || "-")}</td>
                    <td><button class="switch is-on" type="button" data-reference-toggle>Actif</button></td>
                  </tr>
                `;
              }).join("")}
            </tbody>
          </table>
        </div>
      </section>
    `;
    bindReferenceActions();
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
              ${logs.map((log) => `
                <tr>
                  <td>${fmtDateTime(log.at)}</td>
                  <td>${escapeHtml(log.user)}</td>
                  <td>${escapeHtml(log.action)}</td>
                  <td>${escapeHtml(log.detail)}</td>
                  <td>${escapeHtml(log.objectId || "-")}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </section>
    `;
  }

  function renderEntry() {
    const familyOptions = unique([...DATA.families.map((f) => f.name), ...getAllEvents().map((e) => e.family).filter(Boolean)]);
    const sectionOptions = unique([...CHARGING_SECTIONS, ...DISCHARGE_SECTIONS, ...getAllEvents().map((e) => e.sectionKey).filter(Boolean)]);
    const qualityOptions = unique([...Object.keys(DATA.tonnage[0]?.pesage || {}), ...getAllEvents().map((e) => e.quality).filter(Boolean)]);
    const localCount = getLocalEvents().length;

    els.view.innerHTML = `
      <section class="panel">
        <div class="panel-head">
          <h2>Nouvel arrêt</h2>
          <span class="badge cyan">${localCount} saisies locales</span>
        </div>
        <form id="event-form" class="form-grid">
          ${fieldSelect("sectionKey", "S/E", sectionOptions)}
          <label>Sous-équipement<input name="subEquipment" placeholder="PD10, RC134, CA30"></label>
          ${fieldSelect("family", "Famille", familyOptions)}
          <label>Début<input name="start" type="datetime-local" required></label>
          <label>Fin<input name="end" type="datetime-local" required></label>
          ${fieldSelect("quality", "Qualité", qualityOptions)}
          <label>Affectation<input name="assignment" placeholder="Navire ou zone"></label>
          <label>Destination<input name="destination" placeholder="Destination"></label>
          <label>Durée calculée<input id="duration-preview" readonly value="0 h"></label>
          <label class="full">Description<textarea name="description" placeholder="Nature de l'anomalie"></textarea></label>
          <div class="inline-actions full">
            <button class="primary-button" type="submit">Ajouter</button>
            <button class="danger-button" id="clear-local" type="button">Réinitialiser les saisies locales</button>
          </div>
        </form>
      </section>

      <section class="panel">
        <div class="panel-head">
          <h2>Dernières saisies locales</h2>
          <span class="badge">${localCount}</span>
        </div>
        ${renderEventsTable(getLocalEvents().slice().reverse())}
      </section>
    `;

    const form = document.getElementById("event-form");
    const preview = document.getElementById("duration-preview");
    const updatePreview = () => {
      const start = form.elements.start.value;
      const end = form.elements.end.value;
      preview.value = fmtHours(hoursBetweenLocal(start, end));
    };
    form.elements.start.addEventListener("input", updatePreview);
    form.elements.end.addEventListener("input", updatePreview);
    form.addEventListener("submit", handleEventSubmit);
    document.getElementById("clear-local").addEventListener("click", () => {
      saveLocalEvents([]);
      populateFilters();
      render();
    });
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
      <div class="legend-list">
        ${items.map((item, index) => `
          <div class="legend-row">
            <span class="legend-dot" style="background:${chartColor(index)}"></span>
            <strong>${escapeHtml(item.label)}</strong>
            <span>${fmtHours(item.value)} (${fmtPct(item.value / Math.max(total, 1))})</span>
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
    document.querySelectorAll("[data-target-view]").forEach((button) => {
      button.addEventListener("click", () => {
        state.view = button.dataset.targetView;
        document.querySelectorAll(".nav-item").forEach((nav) => nav.classList.toggle("active", nav.dataset.view === state.view));
        render();
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
    const names = [CURRENT_USER.name, "Ahmed Benali", "Youssef El Amrani", "Khadija Saidi", "Système Excel"];
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
    document.querySelectorAll("[data-detail-id]").forEach((button) => {
      button.addEventListener("click", () => {
        state.selectedEventId = button.dataset.detailId;
        state.view = "stopDetail";
        document.querySelectorAll(".nav-item").forEach((nav) => nav.classList.remove("active"));
        render();
      });
    });
    document.querySelectorAll("[data-validate-id]").forEach((button) => {
      button.addEventListener("click", () => {
        updateStopStatus(button.dataset.validateId, "validated", "Arrêt validé après contrôle.");
      });
    });
    document.querySelectorAll("[data-reject-id]").forEach((button) => {
      button.addEventListener("click", () => {
        const comment = prompt("Motif du rejet", "Information à corriger avant exploitation.");
        updateStopStatus(button.dataset.rejectId, "rejected", comment || "Arrêt rejeté.");
      });
    });
  }

  function updateStopStatus(id, status, comment) {
    const overrides = getValidationOverrides();
    overrides[id] = {
      status,
      by: CURRENT_USER.name,
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
    const days = Math.max(getAllDays().length, 1);
    const runningHours = days * 24 * Math.max(CHARGING_SECTIONS.length, 1) - sum(events, "durationHours");
    return events.length ? runningHours / events.length : runningHours;
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
    return Array.from(grouped.values()).sort((a, b) => b.hours - a.hours).slice(0, 120);
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
    const totalStopHours = sum(events, "durationHours");
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
    const exploitationHours = sumByFamily(chargingEvents, ["exploitation"]);
    const maintenanceHours = sumByFamily(chargingEvents, MAINTENANCE_FAMILIES);
    const chargingStopHours = sum(chargingEvents, "durationHours");
    const runningHours = Math.max(chargingAvailableHours - chargingStopHours, 0);
    const cadenceTph = runningHours ? pesageTotal / runningHours : 0;

    return {
      totalStopHours,
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
      externalHours: sumByFamily(chargingEvents, EXTERNAL_FAMILIES),
      trsExploitation: ratio(chargingAvailableHours - exploitationHours, chargingAvailableHours),
      trsMaintenance: ratio(chargingAvailableHours - maintenanceHours, chargingAvailableHours),
      trsGlobal: ratio(chargingAvailableHours - exploitationHours - maintenanceHours, chargingAvailableHours),
      trgGlobal: ratio(runningHours, chargingAvailableHours),
      cadenceTph
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

    const localEvents = getLocalEvents();
    const newEvent = {
      id: `LOCAL-${Date.now()}`,
      row: null,
      declaredBy: CURRENT_USER.name,
      status: "pending",
      createdAt: new Date().toISOString(),
      sectionKey: form.elements.sectionKey.value,
      subEquipment: form.elements.subEquipment.value.trim(),
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
    saveLocalEvents(localEvents);
    addLog("Création", newEvent.id, "Nouvel arrêt créé et envoyé automatiquement au workflow de validation.");
    populateFilters();
    form.reset();
    state.selectedEventId = newEvent.id;
    state.view = "myStops";
    document.querySelectorAll(".nav-item").forEach((nav) => nav.classList.toggle("active", nav.dataset.view === "myStops"));
    render();
  }

  function bindFlowForms() {
    const trainForm = document.getElementById("train-form");
    const shipForm = document.getElementById("ship-form");
    const shipPreview = document.getElementById("ship-duration-preview");

    trainForm.addEventListener("submit", handleTrainSubmit);
    shipForm.addEventListener("submit", handleShipSubmit);
    document.getElementById("clear-local-trains").addEventListener("click", () => {
      saveLocalTrains([]);
      renderFlow();
    });
    document.getElementById("clear-local-ships").addEventListener("click", () => {
      saveLocalShips([]);
      renderFlow();
    });

    const updateShipPreview = () => {
      shipPreview.value = fmtHours(hoursBetweenLocal(shipForm.elements.start.value, shipForm.elements.end.value));
    };
    shipForm.elements.start.addEventListener("input", updateShipPreview);
    shipForm.elements.end.addEventListener("input", updateShipPreview);
  }

  function handleTrainSubmit(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const durationHours = numberFromInput(form.elements.durationHours.value);
    const delayHours = numberFromInput(form.elements.delayHours.value);
    const totalTonnage = numberFromInput(form.elements.totalTonnage.value);
    const affectationHours = durationHours + delayHours;
    const localTrains = getLocalTrains();

    localTrains.push({
      id: `TRAIN-${Date.now()}`,
      day: `${form.elements.day.value}T00:00:00`,
      trains: numberFromInput(form.elements.trains.value),
      wagons: numberFromInput(form.elements.wagons.value),
      durationHours,
      averageHours: durationHours && numberFromInput(form.elements.trains.value) ? durationHours / numberFromInput(form.elements.trains.value) : 0,
      silos: 0,
      tonnageDA: 0,
      tonnageDB: 0,
      tonnageBascule: totalTonnage,
      totalTonnage,
      affectationHours,
      delayHours,
      cadenceTph: durationHours ? totalTonnage / durationHours : 0,
      trsMaintenanceExploit: affectationHours ? durationHours / affectationHours : 1,
      semiWetTrains: 0,
      observation: form.elements.observation.value.trim()
    });

    saveLocalTrains(localTrains);
    renderFlow();
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
    const localShips = getLocalShips();
    localShips.push({
      id: `SHIP-${Date.now()}`,
      number: getAllShips().length + 1,
      berth: form.elements.berth.value.trim(),
      name: form.elements.name.value.trim(),
      quality: form.elements.quality.value,
      ecNumber: form.elements.ecNumber.value.trim(),
      start: `${start}:00`,
      end: `${end}:00`,
      durationHours,
      scaleA: 0,
      scaleB: 0,
      scaleC: 0,
      scaleD: 0,
      bascule,
      connaissement,
      gapRatio: connaissement ? (connaissement - bascule) / connaissement : 0,
      observation: form.elements.observation.value.trim()
    });

    saveLocalShips(localShips);
    renderFlow();
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

  function drawGauge(id, value, label) {
    const canvas = document.getElementById(id);
    if (!canvas) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const cx = w / 2;
    const cy = h * 0.64;
    const radius = Math.min(w, h) * 0.34;
    const start = Math.PI;
    const end = Math.PI * 2;
    const clamped = Math.max(0, Math.min(Number(value) || 0, 1));

    ctx.clearRect(0, 0, w, h);
    ctx.lineWidth = 18;
    ctx.lineCap = "round";
    ctx.beginPath();
    ctx.arc(cx, cy, radius, start, end);
    ctx.strokeStyle = "#e5edf7";
    ctx.stroke();
    ctx.beginPath();
    ctx.arc(cx, cy, radius, start, start + Math.PI * clamped);
    ctx.strokeStyle = clamped >= 0.85 ? "#20b970" : clamped >= 0.75 ? "#f59e0b" : "#dc2626";
    ctx.stroke();
    ctx.fillStyle = "#10284b";
    ctx.font = "800 30px Segoe UI";
    ctx.textAlign = "center";
    ctx.fillText(fmtPct(clamped), cx, cy - 10);
    ctx.fillStyle = "#64748b";
    ctx.font = "13px Segoe UI";
    ctx.fillText(label, cx, cy + 18);
  }

  function drawDonut(id, data, total) {
    const canvas = document.getElementById(id);
    if (!canvas || !data.length) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const cx = w / 2;
    const cy = h / 2;
    const radius = Math.min(w, h) * 0.36;
    let start = -Math.PI / 2;

    ctx.clearRect(0, 0, w, h);
    data.forEach((item, index) => {
      const angle = (item.value / Math.max(total, 1)) * Math.PI * 2;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, radius, start, start + angle);
      ctx.closePath();
      ctx.fillStyle = chartColor(index);
      ctx.fill();
      start += angle;
    });
    ctx.globalCompositeOperation = "destination-out";
    ctx.beginPath();
    ctx.arc(cx, cy, radius * 0.52, 0, Math.PI * 2);
    ctx.fill();
    ctx.globalCompositeOperation = "source-over";
    ctx.fillStyle = "#10284b";
    ctx.font = "700 15px Segoe UI";
    ctx.textAlign = "center";
    ctx.fillText("Top 5", cx, cy - 4);
    ctx.fillStyle = "#64748b";
    ctx.font = "12px Segoe UI";
    ctx.fillText("arrêts", cx, cy + 14);
  }

  function drawDailyTrend(id, rows) {
    const canvas = document.getElementById(id);
    if (!canvas || !rows.length) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const pad = { left: 42, right: 18, top: 18, bottom: 42 };
    const plotW = w - pad.left - pad.right;
    const plotH = h - pad.top - pad.bottom;
    const max = Math.max(...rows.map((row) => row.stopHours), 1);

    ctx.clearRect(0, 0, w, h);
    drawAxis(ctx, pad, w, h);
    ctx.beginPath();
    rows.forEach((row, index) => {
      const x = pad.left + (plotW * index) / Math.max(rows.length - 1, 1);
      const y = pad.top + plotH - (row.stopHours / max) * plotH;
      if (index === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.strokeStyle = "#dc2626";
    ctx.lineWidth = 2.5;
    ctx.stroke();

    rows.forEach((row, index) => {
      const x = pad.left + (plotW * index) / Math.max(rows.length - 1, 1);
      const y = pad.top + plotH - (row.stopHours / max) * plotH;
      ctx.beginPath();
      ctx.arc(x, y, 4, 0, Math.PI * 2);
      ctx.fillStyle = "#ffffff";
      ctx.fill();
      ctx.strokeStyle = "#dc2626";
      ctx.lineWidth = 2;
      ctx.stroke();
      ctx.fillStyle = "#475569";
      ctx.font = "11px Segoe UI";
      ctx.textAlign = "center";
      ctx.fillText(dayLabel(`${row.day}T00:00:00`), x, h - 16);
    });
  }

  function drawCircuitBars(id, rows) {
    const canvas = document.getElementById(id);
    if (!canvas || !rows.length) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const pad = { left: 38, right: 18, top: 18, bottom: 46 };
    const plotW = w - pad.left - pad.right;
    const plotH = h - pad.top - pad.bottom;

    ctx.clearRect(0, 0, w, h);
    drawAxis(ctx, pad, w, h);
    rows.forEach((row, index) => {
      const x = pad.left + (index * plotW) / rows.length + 12;
      const barW = Math.max(28, plotW / rows.length - 24);
      const barH = Math.max(2, row.value * plotH);
      ctx.fillStyle = chartColor(index + 4);
      ctx.fillRect(x, pad.top + plotH - barH, barW, barH);
      ctx.fillStyle = "#10284b";
      ctx.font = "700 12px Segoe UI";
      ctx.textAlign = "center";
      ctx.fillText(fmtPct(row.value), x + barW / 2, pad.top + plotH - barH - 8);
      ctx.fillStyle = "#475569";
      ctx.font = "11px Segoe UI";
      ctx.fillText(row.label, x + barW / 2, h - 16);
    });
  }

  function drawPareto(id, data, total) {
    const canvas = document.getElementById(id);
    if (!canvas) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const pad = { left: 48, right: 24, top: 18, bottom: 72 };
    const plotW = w - pad.left - pad.right;
    const plotH = h - pad.top - pad.bottom;
    const max = Math.max(...data.map((d) => d.value), 1);
    let cumulative = 0;

    ctx.clearRect(0, 0, w, h);
    drawAxis(ctx, pad, w, h);
    data.forEach((item, i) => {
      const x = pad.left + (i * plotW) / data.length + 8;
      const barW = Math.max(12, plotW / data.length - 14);
      const barH = (item.value / max) * plotH;
      ctx.fillStyle = i < 4 ? "#0f766e" : "#7a8d85";
      ctx.fillRect(x, pad.top + plotH - barH, barW, barH);
      ctx.fillStyle = "#3d4641";
      ctx.font = "11px Segoe UI";
      ctx.save();
      ctx.translate(x + 4, h - pad.bottom + 12);
      ctx.rotate(-Math.PI / 4);
      ctx.fillText(truncate(item.label, 18), 0, 0);
      ctx.restore();
      cumulative += item.value;
      const cx = x + barW / 2;
      const cy = pad.top + plotH - (total ? cumulative / total : 0) * plotH;
      if (i === 0) ctx.beginPath();
      ctx.lineTo(cx, cy);
    });
    ctx.strokeStyle = "#b66a0a";
    ctx.lineWidth = 2;
    ctx.stroke();
  }

  function drawBars(id, data, options = {}) {
    const canvas = document.getElementById(id);
    if (!canvas) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const pad = { left: 46, right: 18, top: 18, bottom: 68 };
    const plotW = w - pad.left - pad.right;
    const plotH = h - pad.top - pad.bottom;
    const max = Math.max(...data.map((d) => d.value), 1);
    ctx.clearRect(0, 0, w, h);
    drawAxis(ctx, pad, w, h);
    data.forEach((item, i) => {
      const x = pad.left + (i * plotW) / data.length + 5;
      const barW = Math.max(7, plotW / data.length - 10);
      const barH = (item.value / max) * plotH;
      ctx.fillStyle = options.color || "#0f766e";
      ctx.fillRect(x, pad.top + plotH - barH, barW, barH);
      ctx.fillStyle = "#3d4641";
      ctx.font = "11px Segoe UI";
      ctx.save();
      ctx.translate(x, h - pad.bottom + 14);
      ctx.rotate(-Math.PI / 4);
      ctx.fillText(truncate(item.label, 14), 0, 0);
      ctx.restore();
    });
  }

  function drawLineBars(id, rows) {
    const canvas = document.getElementById(id);
    if (!canvas) return;
    const ctx = setupCanvas(canvas);
    const w = canvas.width / pixelRatio();
    const h = canvas.height / pixelRatio();
    const pad = { left: 50, right: 28, top: 20, bottom: 58 };
    const plotW = w - pad.left - pad.right;
    const plotH = h - pad.top - pad.bottom;
    const max = Math.max(...rows.flatMap((r) => [r.bar, r.line]), 1);
    ctx.clearRect(0, 0, w, h);
    drawAxis(ctx, pad, w, h);
    rows.forEach((row, i) => {
      const x = pad.left + (i * plotW) / rows.length + 3;
      const barW = Math.max(5, plotW / rows.length - 6);
      const barH = (row.bar / max) * plotH;
      ctx.fillStyle = "#0f766e";
      ctx.fillRect(x, pad.top + plotH - barH, barW, barH);
      if (i % 3 === 0) {
        ctx.fillStyle = "#3d4641";
        ctx.font = "11px Segoe UI";
        ctx.fillText(row.label, x - 2, h - pad.bottom + 18);
      }
    });
    ctx.beginPath();
    rows.forEach((row, i) => {
      const x = pad.left + (i * plotW) / rows.length + Math.max(5, plotW / rows.length - 6) / 2;
      const y = pad.top + plotH - (row.line / max) * plotH;
      if (i === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.strokeStyle = "#b66a0a";
    ctx.lineWidth = 2;
    ctx.stroke();
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
