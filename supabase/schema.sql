create extension if not exists "pgcrypto";

create table if not exists public.records (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  row_data jsonb not null default '{}'::jsonb,
  invoice_data jsonb not null default '{}'::jsonb,
  transfer_data jsonb not null default '{}'::jsonb,
  checks_data jsonb not null default '[]'::jsonb,
  source text not null default 'upload',
  invoice_key text not null,
  status text not null default 'Da pagare'
);

create unique index if not exists records_invoice_key_key on public.records (invoice_key);
create index if not exists records_created_at_idx on public.records (created_at);

create table if not exists public.suppliers (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  vat text not null default '',
  iban text not null default '',
  swift text not null default '',
  bank text not null default '',
  currency text not null default 'EUR',
  notes text not null default '',
  updated_at timestamptz not null default now()
);

create unique index if not exists suppliers_name_key on public.suppliers (name);

alter table public.records enable row level security;
alter table public.suppliers enable row level security;

drop policy if exists "records_select_all" on public.records;
create policy "records_select_all"
on public.records
for select
to anon, authenticated
using (true);

drop policy if exists "records_insert_all" on public.records;
create policy "records_insert_all"
on public.records
for insert
to anon, authenticated
with check (true);

drop policy if exists "records_update_all" on public.records;
create policy "records_update_all"
on public.records
for update
to anon, authenticated
using (true)
with check (true);

drop policy if exists "records_delete_all" on public.records;
create policy "records_delete_all"
on public.records
for delete
to anon, authenticated
using (true);

drop policy if exists "suppliers_select_all" on public.suppliers;
create policy "suppliers_select_all"
on public.suppliers
for select
to anon, authenticated
using (true);

drop policy if exists "suppliers_insert_all" on public.suppliers;
create policy "suppliers_insert_all"
on public.suppliers
for insert
to anon, authenticated
with check (true);

drop policy if exists "suppliers_update_all" on public.suppliers;
create policy "suppliers_update_all"
on public.suppliers
for update
to anon, authenticated
using (true)
with check (true);

insert into storage.buckets (id, name, public)
values ('documents', 'documents', false)
on conflict (id) do nothing;

drop policy if exists "documents_select_all" on storage.objects;
create policy "documents_select_all"
on storage.objects
for select
to anon, authenticated
using (bucket_id = 'documents');

drop policy if exists "documents_insert_all" on storage.objects;
create policy "documents_insert_all"
on storage.objects
for insert
to anon, authenticated
with check (bucket_id = 'documents');
