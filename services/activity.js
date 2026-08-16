/* ============================================================
   SERVICE "activity_log" — historique / fil d'activité par dossier
   ============================================================ */

import { supabase } from "../supabaseClient";

const TABLE = "activity_log";

export async function fetchActivity(clientId, { limit = 50 } = {}) {
  const { data, error } = await supabase
    .from(TABLE)
    .select("*")
    .eq("client_id", clientId)
    .order("created_at", { ascending: false })
    .limit(limit);
  if (error) { console.error("Erreur chargement historique :", error.message); return []; }
  return data || [];
}

// À appeler à chaque événement métier significatif : changement de statut,
// création/complétion de tâche, ajout de document, import Excel, etc.
export async function logActivity({ clientId, portefeuilleId, type = "note", message, auteurId, metadata = {} }) {
  const { data, error } = await supabase
    .from(TABLE)
    .insert({ client_id: clientId, portefeuille_id: portefeuilleId || null, type, message, auteur_id: auteurId || null, metadata })
    .select()
    .single();
  if (error) { console.error("Erreur journalisation activité :", error.message); return null; }
  return data;
}

export function subscribeActivity(clientId, onChange) {
  const channel = supabase
    .channel(`activity-${clientId}`)
    .on("postgres_changes", { event: "INSERT", schema: "public", table: TABLE, filter: `client_id=eq.${clientId}` }, onChange)
    .subscribe();
  return () => supabase.removeChannel(channel);
}

/* ============================================================
   Helpers prêts à l'emploi pour les cas d'usage les plus courants
   (évite de ré-écrire le message à chaque fois dans les composants)
   ============================================================ */
export const activityMessages = {
  statutChange: (ancien, nouveau) => `Statut du dossier changé de « ${ancien} » à « ${nouveau} »`,
  tacheCreee: (nom) => `Tâche créée : ${nom}`,
  tacheTerminee: (nom) => `Tâche terminée : ${nom}`,
  documentAjoute: (nom) => `Document ajouté : ${nom}`,
  importExcel: (nbChamps) => `Fiche mise à jour via import Excel (${nbChamps} champ(s))`,
};
