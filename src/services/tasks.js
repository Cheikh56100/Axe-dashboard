/* ============================================================
   SERVICE "tasks" — accès Supabase pour le système de tâches
   ------------------------------------------------------------
   Toute la logique d'accès aux données des tâches vit ici.
   Les composants React n'appellent jamais `supabase.from("tasks")`
   directement : ils passent par ces fonctions.
   ============================================================ */

import { supabase } from "../supabaseClient";

const TABLE = "tasks";

export async function fetchTasks({ portefeuilleId } = {}) {
  let query = supabase.from(TABLE).select("*").order("date_echeance", { ascending: true, nullsFirst: false });
  if (portefeuilleId) query = query.eq("portefeuille_id", portefeuilleId);
  const { data, error } = await query;
  if (error) { console.error("Erreur chargement tâches :", error.message); return []; }
  return data || [];
}

export async function createTask(task) {
  const { data, error } = await supabase.from(TABLE).insert(task).select().single();
  if (error) { console.error("Erreur création tâche :", error.message); return null; }
  return data;
}

export async function updateTask(id, patch) {
  const { data, error } = await supabase.from(TABLE).update(patch).eq("id", id).select().single();
  if (error) { console.error("Erreur mise à jour tâche :", error.message); return null; }
  return data;
}

// Raccourci pour marquer une tâche terminée (renseigne aussi la date de réalisation)
export async function completeTask(id) {
  return updateTask(id, { statut: "termine", date_realisation: new Date().toISOString().slice(0, 10) });
}

export async function deleteTask(id) {
  const { error } = await supabase.from(TABLE).delete().eq("id", id);
  if (error) { console.error("Erreur suppression tâche :", error.message); return false; }
  return true;
}

// Abonnement temps réel : appelle `onChange` à chaque INSERT/UPDATE/DELETE
// Utilisation : useEffect(() => subscribeTasks(reload), [])
export function subscribeTasks(onChange) {
  const channel = supabase
    .channel("tasks-realtime")
    .on("postgres_changes", { event: "*", schema: "public", table: TABLE }, onChange)
    .subscribe();
  return () => supabase.removeChannel(channel);
}
