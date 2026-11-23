-- Execute no SQL Editor do projeto Supabase antes de ativar as rotas de dias.
create extension if not exists pgcrypto;

create table if not exists public.operation_days (
  id uuid primary key default gen_random_uuid(),
  day date not null unique,
  revision integer not null default 0,
  submissions_closed boolean not null default false,
  created_at timestamptz not null default now()
);

create table if not exists public.day_entries (
  id uuid primary key default gen_random_uuid(),
  day_id uuid not null references public.operation_days(id) on delete cascade,
  kind text not null check (kind in ('file', 'text')),
  file_name text,
  object_path text,
  sha256 text,
  byte_size integer,
  manual_text text,
  store text,
  detected_store text,
  detection_source text,
  review_required boolean not null default false,
  preview jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  check ((kind = 'file' and object_path is not null and file_name is not null)
    or (kind = 'text' and manual_text is not null))
);
create index if not exists day_entries_day_id_idx on public.day_entries(day_id, created_at);

create table if not exists public.day_runs (
  id uuid primary key default gen_random_uuid(),
  day_id uuid not null references public.operation_days(id) on delete cascade,
  revision integer not null,
  state text not null check (state in ('processing', 'completed', 'failed')),
  snapshot jsonb not null,
  summary jsonb,
  warnings jsonb,
  object_path text,
  error_message text,
  created_at timestamptz not null default now(),
  finished_at timestamptz
);
create index if not exists day_runs_day_id_idx on public.day_runs(day_id, created_at desc);
create unique index if not exists day_runs_one_processing_idx on public.day_runs(day_id) where state = 'processing';

alter table public.operation_days enable row level security;
alter table public.day_entries enable row level security;
alter table public.day_runs enable row level security;
-- Sem policies: somente a API com service role acessa estas tabelas.
insert into storage.buckets (id, name, public, file_size_limit)
values ('feira-nova-private', 'feira-nova-private', false, 10485760)
on conflict (id) do nothing;
