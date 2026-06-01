-- export_history table for conversor-de-planilhas-fiscal
create table if not exists export_history (
  id text primary key,
  user_id uuid references auth.users(id) on delete cascade,
  filename text not null,
  client_name text not null,
  month_name text not null,
  timestamp text not null,
  result_rows jsonb not null default '[]',
  created_at timestamptz not null default now()
);

-- Enable RLS
alter table export_history enable row level security;

-- Each user sees only their own records
create policy "select_own" on export_history
  for select using (auth.uid() = user_id);

create policy "insert_own" on export_history
  for insert with check (auth.uid() = user_id);

create policy "delete_own" on export_history
  for delete using (auth.uid() = user_id);
