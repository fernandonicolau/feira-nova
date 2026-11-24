create table if not exists public.day_uploads (
  id uuid primary key default gen_random_uuid(),
  day_id uuid not null references public.operation_days(id) on delete cascade,
  client_upload_id text not null,
  file_name text not null,
  object_path text not null,
  sha256 text not null,
  byte_size integer not null check (byte_size > 0),
  created_at timestamptz not null default now(),
  unique (day_id, client_upload_id)
);
create index if not exists day_uploads_day_id_idx on public.day_uploads(day_id);

alter table public.day_entries add column if not exists upload_id uuid references public.day_uploads(id);
alter table public.day_entries add column if not exists client_entry_id text;
alter table public.day_entries add column if not exists state text not null default 'received'
  check (state in ('received', 'needs_review', 'failed'));
alter table public.day_entries add column if not exists deleted_at timestamptz;
alter table public.day_entries add column if not exists replaced_by uuid;
create unique index if not exists day_entries_client_entry_idx
  on public.day_entries(day_id, client_entry_id) where client_entry_id is not null;

create table if not exists public.day_events (
  id bigint generated always as identity primary key,
  day_id uuid not null references public.operation_days(id) on delete cascade,
  entry_id uuid,
  action text not null check (action in ('uploaded', 'text_added', 'removed', 'replaced')),
  details jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);
create index if not exists day_events_day_id_idx on public.day_events(day_id, created_at desc);

alter table public.day_uploads enable row level security;
alter table public.day_events enable row level security;
-- Sem policies: somente a API com service role lê e escreve.

create or replace function public.bump_day_revision()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  update public.operation_days set revision = revision + 1 where id = coalesce(new.day_id, old.day_id);
  return coalesce(new, old);
end;
$$;
drop trigger if exists day_entries_revision_trigger on public.day_entries;
create trigger day_entries_revision_trigger after insert or update or delete on public.day_entries
  for each row execute function public.bump_day_revision();

create or replace function public.audit_day_entry()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  if tg_op = 'INSERT' then
    insert into public.day_events (day_id, entry_id, action, details)
    values (new.day_id, new.id, case when new.kind = 'text' then 'text_added' else 'uploaded' end,
      jsonb_build_object('fileName', new.file_name, 'uploadId', new.upload_id));
  elsif old.deleted_at is null and new.deleted_at is not null then
    insert into public.day_events (day_id, entry_id, action, details)
    values (new.day_id, new.id, case when new.replaced_by is null then 'removed' else 'replaced' end,
      jsonb_build_object('replacedBy', new.replaced_by));
  end if;
  return new;
end;
$$;
drop trigger if exists day_entries_audit_trigger on public.day_entries;
create trigger day_entries_audit_trigger after insert or update on public.day_entries
  for each row execute function public.audit_day_entry();

create or replace function public.daily_storage_bytes()
returns bigint language sql stable security definer set search_path = public as $$
  select coalesce((select sum(byte_size) from public.day_uploads), 0)
       + coalesce((select sum(e.byte_size) from public.day_entries e
          join public.day_uploads u on u.id = e.upload_id
          where e.object_path is distinct from u.object_path), 0)
       + coalesce((select sum((summary->>'archiveBytes')::bigint) from public.day_runs
          where state = 'completed'), 0);
$$;
revoke execute on function public.daily_storage_bytes() from public, anon, authenticated;
grant execute on function public.daily_storage_bytes() to service_role;
