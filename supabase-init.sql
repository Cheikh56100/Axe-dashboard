-- ============================================================
-- INITIALISATION COMPLÈTE — à coller UNE FOIS dans :
-- Supabase Dashboard > SQL Editor > New query > Run
-- (sur le NOUVEAU projet Supabase)
-- ============================================================

-- ---- Table "clients" : une ligne par dossier, id + blob JSON ----
create table if not exists public.clients (
  id text primary key,
  data jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);
alter table public.clients enable row level security;

drop policy if exists "clients_select" on public.clients;
create policy "clients_select" on public.clients for select using (true);
drop policy if exists "clients_insert" on public.clients;
create policy "clients_insert" on public.clients for insert with check (true);
drop policy if exists "clients_update" on public.clients;
create policy "clients_update" on public.clients for update using (true);
drop policy if exists "clients_delete" on public.clients;
create policy "clients_delete" on public.clients for delete using (true);

-- ---- Table "team" : liste des collaborateurs (partagée, temps réel) ----
create table if not exists public.team (
  id text primary key,
  nom text not null,
  color text not null,
  created_at timestamptz not null default now()
);
alter table public.team enable row level security;

drop policy if exists "team_select" on public.team;
create policy "team_select" on public.team for select using (true);
drop policy if exists "team_insert" on public.team;
create policy "team_insert" on public.team for insert with check (true);
drop policy if exists "team_update" on public.team;
create policy "team_update" on public.team for update using (true);
drop policy if exists "team_delete" on public.team;
create policy "team_delete" on public.team for delete using (true);

-- ---- Active le flux Realtime (INSERT/UPDATE/DELETE) sur les deux tables ----
-- Nécessaire pour que les changements d'un poste apparaissent chez les autres
-- sans recharger la page.
alter publication supabase_realtime add table public.clients;
alter publication supabase_realtime add table public.team;

-- ⚠️ Vérifie ENSUITE dans le Dashboard : Database > Replication
-- que "clients" et "team" apparaissent bien cochées. Si l'étape SQL
-- ci-dessus a déjà été effectuée par le Dashboard, cette commande peut
-- renvoyer une erreur "already member of publication" : c'est normal,
-- ignore-la simplement.
