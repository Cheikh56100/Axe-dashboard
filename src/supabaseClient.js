import { createClient } from "@supabase/supabase-js";

// La clé "anon" de Supabase est prévue pour être publique côté client
// (c'est tes règles RLS sur les tables qui protègent réellement les données).
// On la met donc en dur ici pour éviter tout souci de variables
// d'environnement mal configurées sur Netlify/Vercel.
const supabaseUrl = "https://ybewryneaksqhtlagvxk.supabase.co";
const supabaseAnonKey =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InliZXdyeW5lYWtzcWh0bGFndnhrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODY3ODk2OTYsImV4cCI6MjEwMjM2NTY5Nn0.13uLZry9ivPwFZSJy7a362SSvr6U1HIl_WjkbJO93PY";

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
