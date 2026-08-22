-- ============================================================
-- NOVACAB — ACCÈS ORGANISMES SOCIAUX
-- Données très sensibles : URSSAF, Net-entreprise, SYLAE, CIBTP…
-- À exécuter une fois dans Supabase > SQL Editor.
-- ============================================================

create table if not exists public.acces_organismes_sociaux (
  id uuid primary key default gen_random_uuid(),
  portefeuille_id text not null,
  client_id text not null,
  organisme text not null default 'Autre',
  libelle text not null default '',
  identifiant text not null default '',
  secret text not null default '',
  siret text not null default '',
  note text not null default '',
  created_by text,
  updated_by text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_acces_sociaux_portefeuille on public.acces_organismes_sociaux(portefeuille_id);
create index if not exists idx_acces_sociaux_client on public.acces_organismes_sociaux(client_id);

alter table public.acces_organismes_sociaux enable row level security;

-- Tous les utilisateurs actifs du portefeuille peuvent lire ; les écritures sont réservées aux profils habilités.
drop policy if exists "acces_sociaux_select_expert_cdm" on public.acces_organismes_sociaux;
create policy "acces_sociaux_select_expert_cdm"
on public.acces_organismes_sociaux
for select to authenticated
using (
  exists (
    select 1
    from public.team t
    where t.auth_user_id::text = auth.uid()::text
      and t.statut = 'actif'
      and t.portefeuille_id = acces_organismes_sociaux.portefeuille_id
  )
);

-- Même protection pour la création.
drop policy if exists "acces_sociaux_insert_expert_cdm" on public.acces_organismes_sociaux;
create policy "acces_sociaux_insert_expert_cdm"
on public.acces_organismes_sociaux
for insert to authenticated
with check (
  exists (
    select 1
    from public.team t
    where t.auth_user_id::text = auth.uid()::text
      and t.statut = 'actif'
      and t.role in ('admin', 'expert', 'chef_mission', 'gestionnaire_paie')
      and t.portefeuille_id = acces_organismes_sociaux.portefeuille_id
  )
);

-- Modification : impossible de déplacer un accès vers un autre portefeuille.
drop policy if exists "acces_sociaux_update_expert_cdm" on public.acces_organismes_sociaux;
create policy "acces_sociaux_update_expert_cdm"
on public.acces_organismes_sociaux
for update to authenticated
using (
  exists (
    select 1
    from public.team t
    where t.auth_user_id::text = auth.uid()::text
      and t.statut = 'actif'
      and t.role in ('admin', 'expert', 'chef_mission', 'gestionnaire_paie')
      and t.portefeuille_id = acces_organismes_sociaux.portefeuille_id
  )
)
with check (
  exists (
    select 1
    from public.team t
    where t.auth_user_id::text = auth.uid()::text
      and t.statut = 'actif'
      and t.role in ('admin', 'expert', 'chef_mission', 'gestionnaire_paie')
      and t.portefeuille_id = acces_organismes_sociaux.portefeuille_id
  )
);

-- Suppression réservée aux mêmes profils.
drop policy if exists "acces_sociaux_delete_expert_cdm" on public.acces_organismes_sociaux;
create policy "acces_sociaux_delete_expert_cdm"
on public.acces_organismes_sociaux
for delete to authenticated
using (
  exists (
    select 1
    from public.team t
    where t.auth_user_id::text = auth.uid()::text
      and t.statut = 'actif'
      and t.role in ('admin', 'expert', 'chef_mission', 'gestionnaire_paie')
      and t.portefeuille_id = acces_organismes_sociaux.portefeuille_id
  )
);

-- Realtime pour que les changements apparaissent immédiatement chez les Experts/CDM.
alter publication supabase_realtime add table public.acces_organismes_sociaux;
