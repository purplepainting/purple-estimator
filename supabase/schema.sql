-- =====================================================================
-- Purple Estimator schema
-- Apply via Supabase dashboard -> SQL Editor -> paste & run.
-- =====================================================================

-- Singleton config tables (team-wide, one row keyed by id=1) -----------
create table if not exists scope_library (
  id int primary key default 1 check (id = 1),
  value jsonb not null,
  updated_at timestamptz not null default now()
);

create table if not exists catalog_ids (
  id int primary key default 1 check (id = 1),
  value jsonb not null,
  updated_at timestamptz not null default now()
);

create table if not exists tier_multipliers (
  id int primary key default 1 check (id = 1),
  value jsonb not null,
  updated_at timestamptz not null default now()
);

-- T&M catalog (one row per code) ---------------------------------------
create table if not exists catalog_items (
  id text primary key,
  name text not null,
  code text not null,
  code_name text not null,
  unit_cost numeric not null,
  unit_price numeric not null,
  fits_cost_groups text[] not null default '{}',
  created_at timestamptz not null default now()
);

-- Saved job walks ------------------------------------------------------
create table if not exists job_walks (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  title text,
  data jsonb not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- Chat sessions & messages ---------------------------------------------
create table if not exists chat_sessions (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  title text not null default 'New chat',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists chat_messages (
  id uuid primary key default gen_random_uuid(),
  session_id uuid not null references chat_sessions(id) on delete cascade,
  role text not null check (role in ('user','assistant','system')),
  content jsonb not null,
  created_at timestamptz not null default now()
);

create index if not exists chat_messages_session_idx on chat_messages(session_id, created_at);

-- Per-user key-value fallback ------------------------------------------
-- user_id is text (not uuid + FK) so the no-auth testing mode can use the
-- literal 'public-user' identifier. Re-tighten when auth is added back.
create table if not exists kv_store (
  user_id text not null,
  key text not null,
  value jsonb not null,
  updated_at timestamptz not null default now(),
  primary key (user_id, key)
);

-- =====================================================================
-- Row level security
-- Auth gate is DISABLED for testing — every table is open to anon and
-- authenticated. Tighten before going live with real customer data.
-- =====================================================================

alter table scope_library      enable row level security;
alter table catalog_ids        enable row level security;
alter table tier_multipliers   enable row level security;
alter table catalog_items      enable row level security;
alter table job_walks          enable row level security;
alter table chat_sessions      enable row level security;
alter table chat_messages      enable row level security;
alter table kv_store           enable row level security;

create policy "anon+auth rw scope_library"    on scope_library    for all to anon, authenticated using (true) with check (true);
create policy "anon+auth rw catalog_ids"      on catalog_ids      for all to anon, authenticated using (true) with check (true);
create policy "anon+auth rw tier_multipliers" on tier_multipliers for all to anon, authenticated using (true) with check (true);
create policy "anon+auth rw catalog_items"    on catalog_items    for all to anon, authenticated using (true) with check (true);
create policy "anon+auth rw job_walks"        on job_walks        for all to anon, authenticated using (true) with check (true);
create policy "anon+auth rw chat_sessions"    on chat_sessions    for all to anon, authenticated using (true) with check (true);
create policy "anon+auth rw chat_messages"    on chat_messages    for all to anon, authenticated using (true) with check (true);
create policy "anon+auth rw kv_store"         on kv_store         for all to anon, authenticated using (true) with check (true);

-- =====================================================================
-- Seed data
-- =====================================================================

insert into tier_multipliers (id, value) values (1, '{
  "standard":   { "label": "Standard",        "multiplier": 1.00, "color": "#4A90D9", "prepTier": "standard" },
  "production": { "label": "Production",      "multiplier": 0.85, "color": "#7B68EE", "prepTier": "standard" },
  "highend":    { "label": "High-End",        "multiplier": 1.35, "color": "#C8963E", "prepTier": "high_end" },
  "prevailing": { "label": "Prevailing Wage", "multiplier": 1.65, "color": "#E05C5C", "prepTier": "high_end" }
}'::jsonb)
on conflict (id) do nothing;

insert into catalog_ids (id, value) values (1, '{
  "walls":     { "1coat": "22PWT9nDL9HP", "2coats": "22PWiSS293E2", "prime+2": "22PWiSgjqkaq", "code": "22PWmcdKrCnn" },
  "ceilings":  { "1coat": "22PWickkTi46", "2coats": "22PWictZ3V84", "prime+2": "22PWicwN2pxL", "code": "22PWmcdMcqbZ" },
  "baseboard": { "1coat": "22PWiYQt5Pq8", "2coats": "22PWih8XPvPm", "code": "22PWmcdjtJ9C" },
  "doors":     { "1coat": "22PXFmQz3HEd", "2coats": "22PXFmRbHAgn", "prime+2": "22PXFmSG2zeV", "code": "22PWmcdpBtSV" }
}'::jsonb)
on conflict (id) do nothing;

insert into catalog_items (id, name, code, code_name, unit_cost, unit_price, fits_cost_groups) values
  ('22PWTAs6vVPw', 'Time & Materials',                              '1000', 'Interior Walls & Ceilings',     43.33, 65, '{drywall_walls_ceilings}'),
  ('22PWmfVWsyZ9', 'Time & Materials - Doors',                      '2000', 'Doors & Windows',               43.34, 65, '{doors_frames}'),
  ('22PWmfVZLfua', 'Time & Materials - Windows',                    '2000', 'Doors & Windows',               43.34, 65, '{doors_frames}'),
  ('22PWmfVaqmfL', 'Time & Materials - Trim & Beams',               '4000', 'Trim & Beams',                  43.34, 65, '{baseboard_trim}'),
  ('22PWmfVcQeW9', 'Time & Materials - Exterior Walls & Ceilings',  '5000', 'Exterior Walls & Ceilings',     43.34, 65, '{exterior_stucco_siding}'),
  ('22PWmfVec8ks', 'Time & Materials - Exterior Miscellaneous',     '6000', 'Exterior Miscellaneous',        43.34, 65, '{exterior_wood_trim}'),
  ('22PWmfVgD8nW', 'Time & Materials - Millwork & Specialty',       '7000', 'Millwork & Specialty Coatings', 43.34, 65, '{cabinets}'),
  ('22PWmfVhkPTk', 'Time & Materials - Repairs & Cleaning',         '8000', 'Cleaning & Repairs',            43.34, 65, '{universal}'),
  ('22PWmfVpaFLF', 'Time & Materials - Cleaning Services',          '8200', 'Cleaning',                      36.67, 55, '{universal}'),
  ('22PWmfVmyWGZ', 'Time & Materials - Materials & Supplies',       '9000', 'Non-Labor Costs',               43.34, 65, '{universal}')
on conflict (id) do nothing;
