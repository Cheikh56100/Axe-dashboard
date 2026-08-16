/* ============================================================
   RÈGLES DU MOTEUR D'ÉCHÉANCES — paramétrables
   ------------------------------------------------------------
   Modifie ces valeurs pour changer le comportement de TOUT le
   cabinet sans toucher aux composants React.
   ============================================================ */

export const DEFAULT_DEADLINE_RULES = {
  // en dessous de 0 jour restant → "En retard"
  // = 0 jour restant (aujourd'hui) → "Urgent"
  // = 1 jour restant (demain) → "Prioritaire"
  // entre 2 et joursAVenir jours → "À venir"
  // au-delà → "Planifié" (pas encore dans le radar quotidien)
  joursPrioritaire: 1,
  joursAVenir: 14,

  // Régimes TVA (jour légal d'exigibilité par défaut si le dossier n'a
  // pas de jour personnalisé renseigné)
  tvaCa3JourParDefaut: 20,
  tvaCa12JourParDefaut: 3,
  tvaCa12Mois: 4, // Mai (index 0 = janvier)

  // Acomptes IS : 15 mars / juin / sept / déc, 25% de l'IS N-1 chacun
  isAcompteJour: 15,
  isAcomptePourcentage: 0.25,

  // CFE : acompte 50% en juin, solde en décembre
  cfeAcompteJour: 15,
  cfeAcomptePourcentage: 0.5,

  // Acomptes TVA (régime CA12 uniquement) : 55% en juillet, 40% en décembre
  tvaAcompteJour: 15,
  tvaAcompteJuilletPourcentage: 0.55,
  tvaAcompteDecembrePourcentage: 0.40,

  // Bilan : échéance approximative = clôture + N mois si non finalisé
  bilanDelaiMois: 3,
  // AGE/AGO : approbation attendue = clôture + N mois si non tenue
  ageAgoDelaiMois: 6,

  // Seuils de "concernement" (en dessous, la question ne se pose pas)
  seuilIsConcerne: 0,      // IS : concerné dès que nMoins1 > 0 (logique actuelle : voir isIsConcerne)
  seuilCfeConcerne: 3000,  // CFE : concerné si CFE N-1 > 3000€
  seuilTvaAcompteConcerne: 1000, // TVA CA12 : concerné si TVA N-1 > 1000€
};
