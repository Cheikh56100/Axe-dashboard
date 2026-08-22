-- ============================================================
-- NOVACAB — RÔLE GESTIONNAIRE DE PAIE + DROITS ACCÈS ORGANISMES SOCIAUX
-- À exécuter une fois dans Supabase > SQL Editor.
-- ============================================================

-- Le rôle est stocké dans team.role (TEXT). Aucun ENUM bloquant n'est imposé ici,
-- mais on documente les valeurs officielles utilisées par NOVACAB.
-- Valeurs : collaborateur, expert, chef_mission, gestionnaire_paie, admin.

-- Colonnes attendues par l'application si elles n'existent pas encore.
alter table public.team add column if not exists role text default 'collaborateur';
alter table public.team add column if not exists statut text default 'actif';
alter table public.team add column if not exists portefeuille_id text;
alter table public.team add column if not exists auth_user_id uuid;
alter table public.team add column if not exists email text;
alter table public.team add column if not exists telephone text;
alter table public.team add column if not exists cabinet_nom text;

-- ============================================================
-- ACCÈS ORGANISMES SOCIAUX
-- La rubrique et ses données sont consultables par tous les utilisateurs
-- actifs du portefeuille. Les écritures sont réservées à :
-- Admin / Expert / Chef de mission / Gestionnaire de paie.
-- ============================================================

alter table public.acces_organismes_sociaux enable row level security;

drop policy if exists "acces_sociaux_select_tous_portefeuille" on public.acces_organismes_sociaux;
create policy "acces_sociaux_select_tous_portefeuille"
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

drop policy if exists "acces_sociaux_insert_habilites" on public.acces_organismes_sociaux;
create policy "acces_sociaux_insert_habilites"
on public.acces_organismes_sociaux
for insert to authenticated
with check (
  exists (
    select 1
    from public.team t
    where t.auth_user_id::text = auth.uid()::text
      and t.statut = 'actif'
      and t.portefeuille_id = acces_organismes_sociaux.portefeuille_id
      and t.role in ('admin', 'expert', 'chef_mission', 'gestionnaire_paie')
  )
);

drop policy if exists "acces_sociaux_update_habilites" on public.acces_organismes_sociaux;
create policy "acces_sociaux_update_habilites"
on public.acces_organismes_sociaux
for update to authenticated
using (
  exists (
    select 1
    from public.team t
    where t.auth_user_id::text = auth.uid()::text
      and t.statut = 'actif'
      and t.portefeuille_id = acces_organismes_sociaux.portefeuille_id
      and t.role in ('admin', 'expert', 'chef_mission', 'gestionnaire_paie')
  )
)
with check (
  exists (
    select 1
    from public.team t
    where t.auth_user_id::text = auth.uid()::text
      and t.statut = 'actif'
      and t.portefeuille_id = acces_organismes_sociaux.portefeuille_id
      and t.role in ('admin', 'expert', 'chef_mission', 'gestionnaire_paie')
  )
);

drop policy if exists "acces_sociaux_delete_habilites" on public.acces_organismes_sociaux;
create policy "acces_sociaux_delete_habilites"
on public.acces_organismes_sociaux
for delete to authenticated
using (
  exists (
    select 1
    from public.team t
    where t.auth_user_id::text = auth.uid()::text
      and t.statut = 'actif'
      and t.portefeuille_id = acces_organismes_sociaux.portefeuille_id
      and t.role in ('admin', 'expert', 'chef_mission', 'gestionnaire_paie')
  )
);

-- ============================================================
-- NOTE : la création d'un compte Auth doit rester côté Supabase Auth.
-- Une fois le compte créé et lié à team.auth_user_id, l'Admin peut
-- lui attribuer le rôle gestionnaire_paie depuis NOVACAB > Équipe.
-- ============================================================
