import "dotenv/config";
import express from "express";
import cors from "cors";
import jwt from "jsonwebtoken";
import { z } from "zod";

const app = express();
const port = Number(process.env.PORT || 8080);
const jwtSecret = process.env.JWT_SECRET || "trace-port-dev-secret";

app.use(cors());
app.use(express.json({ limit: "2mb" }));

const users = [
  { id: "u1", email: "agent@trace-port.local", password: "demo", name: "Ahmed Benali", role: "Agent" },
  { id: "u2", email: "chef@trace-port.local", password: "demo", name: "Youssef El Amrani", role: "ChefEquipe" },
  { id: "u3", email: "admin@trace-port.local", password: "demo", name: "Israa El Houdzi", role: "Administrateur" }
];

const stops = [];
const logs = [];

const permissions = {
  Agent: ["stops:create", "stops:read"],
  ChefEquipe: ["stops:read", "stops:validate", "reports:read"],
  Administrateur: ["*"]
};

const stopSchema = z.object({
  equipment: z.string().min(1),
  circuit: z.string().min(1),
  stopType: z.string().min(1),
  startedAt: z.string().datetime(),
  endedAt: z.string().datetime(),
  quality: z.string().optional(),
  assignment: z.string().optional(),
  comment: z.string().optional()
});

function sign(user) {
  return jwt.sign({ sub: user.id, name: user.name, role: user.role }, jwtSecret, { expiresIn: "8h" });
}

function auth(req, res, next) {
  const token = req.headers.authorization?.replace("Bearer ", "");
  if (!token) return res.status(401).json({ message: "Token manquant" });
  try {
    req.user = jwt.verify(token, jwtSecret);
    next();
  } catch {
    res.status(401).json({ message: "Token invalide" });
  }
}

function can(permission) {
  return (req, res, next) => {
    const allowed = permissions[req.user.role] || [];
    if (allowed.includes("*") || allowed.includes(permission)) return next();
    return res.status(403).json({ message: "Accès refusé" });
  };
}

function log(user, action, objectType, objectId, detail) {
  logs.unshift({
    id: `log-${Date.now()}`,
    at: new Date().toISOString(),
    user: user?.name || "Système",
    action,
    objectType,
    objectId,
    detail
  });
}

function computeKpis() {
  const totalDowntimeHours = stops.reduce((acc, stop) => acc + stop.durationHours, 0);
  const availableHours = 31 * 4 * 24;
  const maintenanceHours = stops
    .filter((stop) => ["électrique", "instrumentation", "mécanique", "bande"].includes(stop.stopType.toLowerCase()))
    .reduce((acc, stop) => acc + stop.durationHours, 0);
  const exploitationHours = stops
    .filter((stop) => stop.stopType.toLowerCase() === "exploitation")
    .reduce((acc, stop) => acc + stop.durationHours, 0);

  return {
    totalDowntimeHours,
    stopCount: stops.length,
    trsGlobal: availableHours ? (availableHours - maintenanceHours - exploitationHours) / availableHours : 0,
    trsMaintenance: availableHours ? (availableHours - maintenanceHours) / availableHours : 0,
    trsExploitation: availableHours ? (availableHours - exploitationHours) / availableHours : 0,
    mttrHours: stops.length ? totalDowntimeHours / stops.length : 0
  };
}

app.post("/api/auth/login", (req, res) => {
  const user = users.find((item) => item.email === req.body.email && item.password === req.body.password);
  if (!user) return res.status(401).json({ message: "Identifiants invalides" });
  log(user, "login", "user", user.id, "Connexion utilisateur");
  res.json({ token: sign(user), user: { id: user.id, name: user.name, role: user.role } });
});

app.get("/api/stops", auth, can("stops:read"), (req, res) => {
  res.json(stops);
});

app.post("/api/stops", auth, can("stops:create"), (req, res) => {
  const payload = stopSchema.parse(req.body);
  const durationHours = Math.max(0, (new Date(payload.endedAt) - new Date(payload.startedAt)) / 36e5);
  const stop = {
    id: `AR-${new Date().getFullYear()}-${String(stops.length + 1).padStart(4, "0")}`,
    ...payload,
    durationHours,
    status: "pending",
    declaredBy: req.user.name,
    createdAt: new Date().toISOString()
  };
  stops.push(stop);
  log(req.user, "create", "stop", stop.id, "Nouvel arrêt saisi");
  res.status(201).json(stop);
});

app.post("/api/stops/:id/validate", auth, can("stops:validate"), (req, res) => {
  const stop = stops.find((item) => item.id === req.params.id);
  if (!stop) return res.status(404).json({ message: "Arrêt introuvable" });
  stop.status = "validated";
  stop.validatedBy = req.user.name;
  stop.validatedAt = new Date().toISOString();
  stop.validationComment = req.body.comment || "Validé";
  log(req.user, "validate", "stop", stop.id, stop.validationComment);
  res.json(stop);
});

app.post("/api/stops/:id/reject", auth, can("stops:validate"), (req, res) => {
  const stop = stops.find((item) => item.id === req.params.id);
  if (!stop) return res.status(404).json({ message: "Arrêt introuvable" });
  stop.status = "rejected";
  stop.validatedBy = req.user.name;
  stop.validatedAt = new Date().toISOString();
  stop.validationComment = req.body.comment || "Rejeté";
  log(req.user, "reject", "stop", stop.id, stop.validationComment);
  res.json(stop);
});

app.get("/api/kpis", auth, (req, res) => {
  res.json(computeKpis());
});

app.get("/api/reports/:type", auth, (req, res) => {
  log(req.user, "export", "report", req.params.type, "Génération rapport");
  res.json({ type: req.params.type, generatedAt: new Date().toISOString(), kpis: computeKpis(), stops });
});

app.get("/api/logs", auth, (req, res) => {
  res.json(logs);
});

app.listen(port, () => {
  console.log(`TRACE-PORT API listening on http://localhost:${port}`);
});
