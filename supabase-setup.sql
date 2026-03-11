create extension if not exists "pgcrypto";

create table if not exists public.cycle_entries (
  id text primary key,
  user_id uuid not null references auth.users(id) on delete cascade,
  start_date date not null,
  end_date date not null,
  flow text not null default 'Medium',
  mood text not null default 'Balanced',
  symptoms text[] not null default '{}',
  notes text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create or replace function public.set_cycle_entries_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists cycle_entries_set_updated_at on public.cycle_entries;
create trigger cycle_entries_set_updated_at
before update on public.cycle_entries
for each row
execute function public.set_cycle_entries_updated_at();

alter table public.cycle_entries enable row level security;

create policy "Users can read their own cycle entries"
on public.cycle_entries
for select
using (auth.uid() = user_id);

create policy "Users can insert their own cycle entries"
on public.cycle_entries
for insert
with check (auth.uid() = user_id);

create policy "Users can update their own cycle entries"
on public.cycle_entries
for update
using (auth.uid() = user_id)
with check (auth.uid() = user_id);

create policy "Users can delete their own cycle entries"
on public.cycle_entries
for delete
using (auth.uid() = user_id);
