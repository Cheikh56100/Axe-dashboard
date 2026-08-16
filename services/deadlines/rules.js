/* ============================================================
   RÈGLES DU MOTEUR D'ÉCHÉANCES — paramétrables
   ------------------------------------------------------------
   Modifie ces valeurs pour changer le comportement de TOUT le
   cabinet sans toucher aux composants React.
   ============================================================ */

export const DEFAULT_DEADLINE_RULES = {
  joursPrioritaire: 1,
  joursAVenir: 14,

  tvaCa3JourParDefaut: 20,
  tvaCa12JourParDefaut: 3,
  tvaCa12Mois: 4, // Mai (index 0 = janvier)

  isAcompteJour: 15,
  isAcomptePourcentage: 0.25,

  cfeAcompteJour: 15,
  cfeAcomptePourcentage: 0.5,

  tvaAcompteJour: 15,
  tvaAcompteJuilletPourcentage: 0.55,
  tvaAcompteDecembrePourcentage: 0.40,

  bilanDelaiMois: 3,
  ageAgoDelaiMois: 6,

  seuilIsConcerne: 0,
  seuilCfeConcerne: 3000,
  seuilTvaAcompteConcerne: 1000,
};
