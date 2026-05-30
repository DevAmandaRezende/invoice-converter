-- export_history table for conversor-de-planilhas-fiscal
create table if not exists export_history (
  id text primary key,
  filename text not null,
  client_name text not null,
  month_name text not null,
  timestamp text not null,
  result_rows jsonb not null default '[]',
  created_at timestamptz not null default now()
);

-- Disable RLS (internal tool, no auth)
alter table export_history disable row level security;
