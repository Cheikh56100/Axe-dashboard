/* ============================================================
   MOTEUR D'ANOMALIES — services/anomalies
   ------------------------------------------------------------
   Détecte les incohérences ou données manquantes qui empêchent
   le pilotage correct d'un dossier. Calculé à la volée (pas de
   table dédiée : ce n'est pas un historique, c'est un diagnostic
   instantané de l'état des données).

   Chaque règle est une petite fonction pure (client) => anomalie|null,
   ce qui permet d'en ajouter facilement sans toucher au reste.
   ============================================================ */

import { STATUS } from "../constants/pilotage";

const RULES = [
  {
    code: "siren_manquant",
    gravite: "haute",
    test: (c) => !c.siren,
    message: () => "SIREN non renseigné",
  },
  {
    code: "regime_tva_manquant",
    gravite: "haute",
    test: (c) => c.statutDossier !== "inactif" && !c.tvaRegime,
    message: () => "Régime de TVA non renseigné pour un dossier actif",
  },
  {
    code: "exigibilite_tva_manquante",
    gravite: "moyenne",
    test: (c) => !!c.tvaRegime && !c.tvaExig,
    message: () => "Jour d'exigibilité TVA non renseigné",
  },
  {
    code: "responsable_manquant",
    gravite: "moyenne",
    test: (c) => c.statutDossier !== "inactif" && !c.collab && !c.expert,
    message: () => "Aucun collaborateur ni expert affecté au dossier",
  },
  {
    code: "cloture_manquante",
    gravite: "moyenne",
    test: (c) => c.statutDossier !== "inactif" && !c.dateCloture,
    message: () => "Date de clôture d'exercice non renseignée",
  },
  {
    code: "bilan_tres_en_retard",
    gravite: "haute",
    test: (c) => {
      if (!c.dateCloture || !c.bilan?.nonFinalise) return false;
      const [y, m, d] = c.dateCloture.split("-").map(Number);
      const limite = new Date(y, m - 1 + 6, d); // clôture + 6 mois = alerte forte
      return new Date() > limite;
    },
    message: () => "Bilan non finalisé plus de 6 mois après la clôture",
  },
];

// Analyse un client et retourne la liste de ses anomalies
export function detectClientAnomalies(client) {
  return RULES
    .filter((rule) => rule.test(client))
    .map((rule) => ({
      id: `${client.id}-${rule.code}`,
      clientId: client.id,
      clientNom: client.nom,
      code: rule.code,
      gravite: rule.gravite, // 'haute' | 'moyenne' | 'basse'
      message: rule.message(client),
      status: rule.gravite === "haute" ? STATUS.EN_RETARD : STATUS.A_TRAITER,
    }));
}

// Analyse tout le cabinet
export function detectAllAnomalies(clients = []) {
  return clients.flatMap(detectClientAnomalies);
}

export function countAnomalies(clients = []) {
  return detectAllAnomalies(clients).length;
}
