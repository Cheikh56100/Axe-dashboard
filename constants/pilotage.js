/* ============================================================
   SYSTÈME UNIFIÉ DE STATUTS / PRIORITÉS / COULEURS
   ------------------------------------------------------------
   Un seul endroit pour définir "qu'est-ce que ça veut dire d'être vert,
   orange, rouge, indigo ou gris" dans toute l'application.
   Règle d'or : la couleur seule ne suffit jamais → chaque statut a
   toujours un libellé explicite associé (accessibilité + clarté).
   ============================================================ */

// Palette (reprend les couleurs déjà utilisées dans App.jsx -> objet T,
// pour ne rien casser visuellement).
export const PILOTAGE_COLORS = {
  green:  { text: "#16A34A", bg: "#DCFCE7", dot: "#16A34A" },
  orange: { text: "#D97706", bg: "#FEF3C7", dot: "#D97706" },
  red:    { text: "#DC2626", bg: "#FEE2E2", dot: "#DC2626" },
  indigo: { text: "#4F46E5", bg: "#EEF2FF", dot: "#4F46E5" },
  gray:   { text: "#94A3B8", bg: "#F1F5F9", dot: "#94A3B8" },
};

// Statuts génériques (dossiers, échéances, missions, TVA effective, etc.)
export const STATUS = {
  OK:              { code: "ok",              label: "OK",              color: "green"  },
  A_TRAITER:       { code: "a_traiter",        label: "À traiter",       color: "orange" },
  EN_RETARD:       { code: "en_retard",        label: "En retard",       color: "red"    },
  EN_COURS:        { code: "en_cours",         label: "En cours",        color: "indigo" },
  NON_APPLICABLE:  { code: "non_applicable",   label: "Non applicable",  color: "gray"   },
  BLOQUE:          { code: "bloque",           label: "Bloqué",          color: "red"    },
};

export const STATUS_BY_CODE = Object.fromEntries(
  Object.values(STATUS).map((s) => [s.code, s])
);

export function getStatus(code) {
  return STATUS_BY_CODE[code] || STATUS.NON_APPLICABLE;
}

export function statusColorSet(codeOrStatus) {
  const s = typeof codeOrStatus === "string" ? getStatus(codeOrStatus) : codeOrStatus;
  return PILOTAGE_COLORS[s?.color] || PILOTAGE_COLORS.gray;
}

// ---- Tâches ----
export const TASK_STATUTS = [
  { code: "a_faire",  label: "À faire",  color: "gray"   },
  { code: "en_cours", label: "En cours", color: "indigo" },
  { code: "termine",  label: "Terminé",  color: "green"  },
  { code: "bloque",   label: "Bloqué",   color: "red"    },
];
export const TASK_STATUT_BY_CODE = Object.fromEntries(TASK_STATUTS.map((s) => [s.code, s]));

export const TASK_PRIORITES = [
  { code: "faible",  label: "Faible",  weight: 0 },
  { code: "normale", label: "Normale", weight: 1 },
  { code: "haute",   label: "Haute",   weight: 2 },
  { code: "urgente", label: "Urgente", weight: 3 },
];
export const TASK_PRIORITE_BY_CODE = Object.fromEntries(TASK_PRIORITES.map((p) => [p.code, p]));

// Ordre de tri par défaut : bloqué/urgent d'abord, terminé en dernier
export function taskSortWeight(task) {
  const statutWeight = { bloque: 0, en_cours: 1, a_faire: 2, termine: 3 }[task.statut] ?? 2;
  const prioriteWeight = 3 - (TASK_PRIORITE_BY_CODE[task.priorite]?.weight ?? 1);
  return statutWeight * 10 + prioriteWeight;
}
