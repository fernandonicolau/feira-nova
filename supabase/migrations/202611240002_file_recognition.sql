alter table public.day_entries add column if not exists layout jsonb;
alter table public.day_entries add column if not exists detection_confidence numeric(4,3);
alter table public.day_entries add column if not exists reviewed_at timestamptz;
create index if not exists day_entries_needs_review_idx on public.day_entries(day_id)
  where review_required = true and deleted_at is null;

alter table public.day_events drop constraint if exists day_events_action_check;
alter table public.day_events add constraint day_events_action_check
  check (action in ('uploaded', 'text_added', 'removed', 'replaced', 'reviewed'));

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
  elsif old.reviewed_at is distinct from new.reviewed_at and new.reviewed_at is not null then
    insert into public.day_events (day_id, entry_id, action, details)
    values (new.day_id, new.id, 'reviewed', jsonb_build_object('store', new.store, 'layout', new.layout));
  end if;
  return new;
end;
$$;
