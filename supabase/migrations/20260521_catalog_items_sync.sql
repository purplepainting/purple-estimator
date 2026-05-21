-- Migration: extend catalog_items to mirror the JobTread org catalog.
-- Applied to project aoykceutyrrzqmtbgmfp (purple-estimator) on 2026-05-21.
--
-- Context: catalog_items previously held only a 10-row T&M seed (id, name,
-- code, code_name, unit_cost, unit_price, fits_cost_groups) with code_name /
-- unit_cost / unit_price all NOT NULL. The full JT catalog (~491 items) has
-- unpriced placeholder items and items without a code_name, so those legacy
-- NOT NULLs are relaxed. New columns capture the link-critical ids (code_id =
-- the real cost-code JobTread id, fixing the T&M costCodeId bug) plus cost
-- type, unit, and the three catalog custom fields.

-- 1. Additive columns (sync target shape).
ALTER TABLE public.catalog_items
  ADD COLUMN IF NOT EXISTS code_id        text,   -- cost code JT id (T&M fix; was missing)
  ADD COLUMN IF NOT EXISTS cost_type_id   text,
  ADD COLUMN IF NOT EXISTS cost_type_name text,
  ADD COLUMN IF NOT EXISTS unit_id        text,
  ADD COLUMN IF NOT EXISTS unit_name      text,
  ADD COLUMN IF NOT EXISTS substrate      text,   -- custom field 22PWTCViXMi8 (nullable; T&M has none)
  ADD COLUMN IF NOT EXISTS condition      text,   -- custom field 22PWYi74BtMp
  ADD COLUMN IF NOT EXISTS coats          text,   -- custom field 22PWjDaWGG46
  ADD COLUMN IF NOT EXISTS kind           text NOT NULL DEFAULT 'catalog', -- 'tm' | 'catalog'
  ADD COLUMN IF NOT EXISTS synced_at      timestamptz;

-- 2. Name is the stable identity across JT id changes (names are unique in JT).
--    Sync upserts ON CONFLICT (name); orphan cleanup deletes rows with an older
--    synced_at than the current run.
CREATE UNIQUE INDEX IF NOT EXISTS catalog_items_name_key ON public.catalog_items (name);

-- 3. Relax legacy NOT NULLs that the real catalog violates.
ALTER TABLE public.catalog_items ALTER COLUMN code_name  DROP NOT NULL;
ALTER TABLE public.catalog_items ALTER COLUMN unit_cost  DROP NOT NULL;
ALTER TABLE public.catalog_items ALTER COLUMN unit_price DROP NOT NULL;
