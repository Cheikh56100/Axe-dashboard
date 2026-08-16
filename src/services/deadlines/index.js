/* ============================================================
   MOTEUR D'ÉCHÉANCES — services/deadlines
   ------------------------------------------------------------
   Point d'entrée unique pour tout calcul de date/statut d'échéance
   dans l'application. Ne dépend d'aucun composant React : peut être
   testé isolément et réutilisé partout (Dashboard, Vue Direction,
   Mes tâches, fiche client, notifications...).

   Reprend et généralise la logique qui existait dans
   computeFiscalEvents() / effectiveTvaStatus() / taskBucket() de
   App.jsx, en la rendant paramétrable (voir rules.js) et en lui
   donnant un statut UNIFIÉ (voir constants/pilotage.js) au lieu de
   "tone" ad-hoc.
   ============================================================ */

import { DEFAULT_DEADLINE_RULES } from "./rules";
import { STATUS } from "../../constants/pilotage";

export const MOIS_ORDER = [
  "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
  "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
];

// ------------------------------------------------------------
// Utilitaires de date (purs, sans dépendance)
// ------------------------------------------------------------
export function todayAtMidnight(now = new Date()) {
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

export function diffInDays(date, now = new Date()) {
  const today = todayAtMidnight(now);
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.round((d - today) / 86400000);
}

export function addMonthsISO(iso, months) {
  if (!iso) return null;
  const [y, m, d] = iso.split("-").map(Number);
  const dt = new Date(y, m - 1 + months, d);
  return dt.toISOString().slice(0, 10);
}

// ------------------------------------------------------------
// Statut d'une échéance à partir de son nombre de jours restants
// ------------------------------------------------------------
// diffDays < 0                                → EN_RETARD
// diffDays === 0                              → statut "urgent" (voir DEADLINE_URGENCY.URGENT)
// diffDays === 1..joursPrioritaire            → "prioritaire"
// diffDays <= joursAVenir                     → "a_venir"
// diffDays > joursAVenir                      → "planifie"
export const DEADLINE_URGENCY = {
  EN_RETARD:   { code: "en_retard",   label: "En retard",    color: "red"    },
  URGENT:      { code: "urgent",      label: "Urgent",       color: "red"    },
  PRIORITAIRE: { code: "prioritaire", label: "Prioritaire",  color: "orange" },
  A_VENIR:     { code: "a_venir",     label: "À venir",      color: "orange" },
  PLANIFIE:    { code: "planifie",    label: "Planifié",     color: "indigo" },
};

export function getDeadlineUrgency(date, rules = DEFAULT_DEADLINE_RULES, now = new Date()) {
  const diffDays = diffInDays(date, now);
  let urgency;
  if (diffDays < 0) urgency = DEADLINE_URGENCY.EN_RETARD;
  else if (diffDays === 0) urgency = DEADLINE_URGENCY.URGENT;
  else if (diffDays <= rules.joursPrioritaire) urgency = DEADLINE_URGENCY.PRIORITAIRE;
  else if (diffDays <= rules.joursAVenir) urgency = DEADLINE_URGENCY.A_VENIR;
  else urgency = DEADLINE_URGENCY.PLANIFIE;
  return { ...urgency, diffDays };
}

// Regroupe une liste d'échéances (ou de tâches) par "bucket" temporel,
// pour la page "Mes tâches" (Aujourd'hui / En retard / Cette semaine / À venir).
export function bucketize(items, getDate, now = new Date()) {
  const today = todayAtMidnight(now);
  const weekStart = new Date(today);
  weekStart.setDate(today.getDate() - ((today.getDay() + 6) % 7)); // lundi
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);

  const buckets = { retard: [], aujourdhui: [], semaine: [], avenir: [] };
  items.forEach((item) => {
    const date = getDate(item);
    if (!date) return;
    const diffDays = diffInDays(date, now);
    if (diffDays < 0) buckets.retard.push(item);
    else if (diffDays === 0) buckets.aujourdhui.push(item);
    else if (date >= weekStart && date <= weekEnd) buckets.semaine.push(item);
    else buckets.avenir.push(item);
  });
  return buckets;
}

export const BUCKET_LABELS = {
  retard: "En retard",
  aujourdhui: "Aujourd'hui",
  semaine: "Cette semaine",
  avenir: "À venir",
};

// ------------------------------------------------------------
// Concernement (portage direct des règles existantes)
// ------------------------------------------------------------
export function isIsConcerne(client) {
  return !!(parseFloat(client?.is?.nMoins1) > 0);
}
export function isCfeConcerne(client, rules = DEFAULT_DEADLINE_RULES) {
  const n = parseFloat(client?.cfe?.nMoins1);
  return !!n && n > rules.seuilCfeConcerne;
}
export function isTvaAcompteConcerne(client, rules = DEFAULT_DEADLINE_RULES) {
  if (client?.tvaRegime !== "CA12") return false;
  const n = parseFloat(client?.tvaAcompte?.nMoins1);
  return !!n && n > rules.seuilTvaAcompteConcerne;
}

// ------------------------------------------------------------
// Statut TVA effectif (portage direct de effectiveTvaStatus)
// ------------------------------------------------------------
export function effectiveTvaStatus(client, moisKey, rules = DEFAULT_DEADLINE_RULES, now = new Date()) {
  const manual = (client?.tvaMois?.[moisKey] || "").toUpperCase();
  if (manual === "OK" || manual === "FAIT" || manual === "NA") return manual;

  if (client?.tvaRegime === "CA12") {
    if (moisKey !== "Mai") return "";
    const exig = parseInt(client.tvaExig, 10) || rules.tvaCa12JourParDefaut;
    const deadline = new Date(now.getFullYear(), rules.tvaCa12Mois, exig, 23, 59, 59);
    return deadline.getTime() < now.getTime() ? "RETARD" : "";
  }

  const exig = parseInt(client?.tvaExig, 10);
  if (!exig) return "";
  const monthIdx = MOIS_ORDER.indexOf(moisKey);
  const deadline = new Date(now.getFullYear(), monthIdx + 1, exig, 23, 59, 59);
  return deadline.getTime() < now.getTime() ? "RETARD" : "";
}

// ------------------------------------------------------------
// Construction de la liste d'échéances fiscales/juridiques d'UN client
// ------------------------------------------------------------
// Retourne des objets { id, clientId, category, label, date, montant,
// status: STATUS.*, urgency: DEADLINE_URGENCY.* } — prêts à afficher.
export function buildClientDeadlines(client, rules = DEFAULT_DEADLINE_RULES, now = new Date()) {
  const events = [];
  const year = now.getFullYear();
  const push = (partial) => {
    const urgency = getDeadlineUrgency(partial.date, rules, now);
    const status =
      urgency.code === "en_retard" ? STATUS.EN_RETARD :
      urgency.code === "urgent" || urgency.code === "prioritaire" ? STATUS.A_TRAITER :
      STATUS.EN_COURS;
    events.push({ clientId: client.id, ...partial, urgency, status });
  };

  // TVA CA3 — déclaration du mois M-1, exigible ce mois-ci
  if (client.tvaRegime === "CA3" && client.tvaExig) {
    const monthIdx = now.getMonth();
    const declaredMonthIdx = (monthIdx - 1 + 12) % 12;
    const statut = effectiveTvaStatus(client, MOIS_ORDER[declaredMonthIdx], rules, now);
    if (statut !== "OK" && statut !== "NA") {
      push({
        id: `${client.id}-tva-${declaredMonthIdx}`, category: "TVA",
        label: `TVA ${MOIS_ORDER[declaredMonthIdx]}`,
        date: new Date(year, monthIdx, parseInt(client.tvaExig, 10) || rules.tvaCa3JourParDefaut),
      });
    }
  }
  // TVA CA12 — échéance annuelle en mai N+1
  if (client.tvaRegime === "CA12") {
    const statut = effectiveTvaStatus(client, "Mai", rules, now);
    if (statut !== "OK" && statut !== "NA") {
      push({
        id: `${client.id}-tva-ca12`, category: "TVA", label: "TVA annuelle (CA12)",
        date: new Date(year, rules.tvaCa12Mois, parseInt(client.tvaExig, 10) || rules.tvaCa12JourParDefaut),
      });
    }
  }
  // IS — 4 acomptes trimestriels sur l'IS N-1
  if (isIsConcerne(client)) {
    const montant = Math.round((parseFloat(client.is.nMoins1) || 0) * rules.isAcomptePourcentage);
    [["mars", 2], ["juin", 5], ["sept", 8], ["dec", 11]].forEach(([key, m]) => {
      if (!client.is[key]) {
        push({
          id: `${client.id}-is-${key}`, category: "IS",
          label: `Acompte IS (${montant.toLocaleString("fr-FR")} €)`,
          date: new Date(year, m, rules.isAcompteJour), montant,
        });
      }
    });
  }
  // CFE — acompte juin + solde décembre sur la CFE N-1
  if (isCfeConcerne(client, rules)) {
    const montant = Math.round((parseFloat(client.cfe.nMoins1) || 0) * rules.cfeAcomptePourcentage);
    [["juin", 5, "acompte"], ["dec", 11, "solde"]].forEach(([key, m, label]) => {
      if (!client.cfe[key]) {
        push({
          id: `${client.id}-cfe-${key}`, category: "CFE",
          label: `CFE — ${label} (${montant.toLocaleString("fr-FR")} €)`,
          date: new Date(year, m, rules.cfeAcompteJour), montant,
        });
      }
    });
  }
  // Acomptes TVA (régime CA12) — 55% juillet / 40% décembre sur la TVA N-1
  if (isTvaAcompteConcerne(client, rules)) {
    const nMoins1 = parseFloat(client.tvaAcompte?.nMoins1) || 0;
    [
      ["juillet", 6, Math.round(nMoins1 * rules.tvaAcompteJuilletPourcentage)],
      ["decembre", 11, Math.round(nMoins1 * rules.tvaAcompteDecembrePourcentage)],
    ].forEach(([key, m, montant]) => {
      if (!client.tvaAcompte?.[key]) {
        push({
          id: `${client.id}-tvaacompte-${key}`, category: "TVA",
          label: `Acompte TVA (${montant.toLocaleString("fr-FR")} €)`,
          date: new Date(year, m, rules.tvaAcompteJour), montant,
        });
      }
    });
  }
  // Bilan — clôture + délai si non finalisé
  if (client.dateCloture && client.bilan?.nonFinalise) {
    const echeanceISO = addMonthsISO(client.dateCloture, rules.bilanDelaiMois);
    const [by, bm, bd] = echeanceISO.split("-").map(Number);
    push({
      id: `${client.id}-bilan`, category: "Bilan", label: "Dépôt du bilan",
      date: new Date(by, bm - 1, bd),
    });
  }
  // AGE/AGO — clôture + délai si non tenue
  if (client.dateCloture) {
    const latestYear = Object.keys(client.ageAgoHistory || {}).sort((a, b) => b - a)[0];
    const y = latestYear ? client.ageAgoHistory[latestYear] : null;
    if (y && !y.ago) {
      const echeanceISO = addMonthsISO(client.dateCloture, rules.ageAgoDelaiMois);
      const [ay, am, ad] = echeanceISO.split("-").map(Number);
      push({
        id: `${client.id}-ago-${latestYear}`, category: "AGO",
        label: `Approbation des comptes ${latestYear}`, date: new Date(ay, am - 1, ad),
      });
    }
  }

  return events;
}

// ------------------------------------------------------------
// Transforme une TÂCHE (table "tasks") en objet "échéance" homogène,
// pour pouvoir la mélanger avec les échéances fiscales dans les mêmes vues.
// ------------------------------------------------------------
export function buildTaskDeadline(task, rules = DEFAULT_DEADLINE_RULES, now = new Date()) {
  if (!task.date_echeance) return null;
  const [y, m, d] = task.date_echeance.split("-").map(Number);
  const date = new Date(y, m - 1, d);
  const urgency = getDeadlineUrgency(date, rules, now);
  const status =
    task.statut === "termine" ? STATUS.OK :
    task.statut === "bloque" ? STATUS.BLOQUE :
    urgency.code === "en_retard" ? STATUS.EN_RETARD :
    task.statut === "en_cours" ? STATUS.EN_COURS :
    STATUS.A_TRAITER;
  return {
    id: `task-${task.id}`, taskId: task.id, clientId: task.client_id,
    category: "Tâche", label: task.nom, date, urgency, status,
  };
}

// ------------------------------------------------------------
// Agrège TOUTES les échéances (fiscales + tâches) pour l'ensemble
// du cabinet, ou pour un client donné. Sert de source unique pour :
// Dashboard, Vue Direction, Mes tâches, fiche client (onglet Échéances).
// ------------------------------------------------------------
export function aggregateDeadlines({ clients = [], tasks = [], rules = DEFAULT_DEADLINE_RULES, now = new Date() }) {
  const fiscal = clients.flatMap((c) => buildClientDeadlines(c, rules, now));
  const fromTasks = tasks.map((t) => buildTaskDeadline(t, rules, now)).filter(Boolean);
  return [...fiscal, ...fromTasks].sort((a, b) => a.date - b.date);
}
