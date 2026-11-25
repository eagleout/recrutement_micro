-- Recrutement micro-crèches : schéma Supabase
-- À exécuter une fois dans Supabase › SQL Editor.
-- Remplacez l'adresse de la dernière ligne par la vôtre, puis ajoutez les membres de l'équipe.

-- 1. Membres autorisés (une ligne par adresse e-mail)
create table if not exists public.membres (
  email text primary key,
  ajoute_le timestamptz not null default now()
);

-- 2. Données de l'application : un document JSON par candidat, plus les réglages
create table if not exists public.records (
  collection text not null,              -- 'candidats' ou 'config'
  id         text not null,
  data       jsonb not null,
  updated_at timestamptz not null default now(),
  updated_by uuid default auth.uid(),
  primary key (collection, id)
);

-- 3. Règles d'accès : seuls les membres connectés lisent et écrivent
alter table public.membres enable row level security;
alter table public.records enable row level security;

create or replace function public.est_membre() returns boolean
language sql stable security definer set search_path = public as $$
  select exists (
    select 1 from public.membres m
    where lower(m.email) = lower(coalesce(auth.jwt() ->> 'email', ''))
  );
$$;

drop policy if exists "membre voit sa ligne" on public.membres;
create policy "membre voit sa ligne" on public.membres
  for select to authenticated
  using (lower(email) = lower(coalesce(auth.jwt() ->> 'email', '')));

drop policy if exists "membres lisent" on public.records;
create policy "membres lisent" on public.records
  for select to authenticated using (public.est_membre());

drop policy if exists "membres ajoutent" on public.records;
create policy "membres ajoutent" on public.records
  for insert to authenticated with check (public.est_membre());

drop policy if exists "membres modifient" on public.records;
create policy "membres modifient" on public.records
  for update to authenticated using (public.est_membre()) with check (public.est_membre());

drop policy if exists "membres suppriment" on public.records;
create policy "membres suppriment" on public.records
  for delete to authenticated using (public.est_membre());

-- 4. Synchronisation en temps réel entre appareils
alter table public.records replica identity full;
do $$ begin
  alter publication supabase_realtime add table public.records;
exception when duplicate_object then null; end $$;

-- 5. Premier membre (remplacez par votre adresse), puis une ligne par personne de l'équipe
insert into public.membres (email) values ('votre.email@exemple.fr')
on conflict (email) do nothing;
