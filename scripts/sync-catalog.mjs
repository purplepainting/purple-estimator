#!/usr/bin/env node
/**
 * sync-catalog.mjs — Pull the full JobTread org catalog into Supabase
 * `catalog_items`. This is the source-of-truth sync: JobTread is authoritative,
 * Supabase is the fast queryable mirror the app reads. Re-run whenever the JT
 * catalog changes ("updates when JT updates" — manual for now; this is the seed
 * of the future AWS scheduled/webhook sync).
 *
 * WHY THIS EXISTS / WHAT IT GUARANTEES
 *  - Stores the real JT cost-item id (= organizationCostItemId) so budget cost
 *    items link correctly on create. The hardcoded TM_CATALOG had a STALE id
 *    (22PWTAs6vVPw, deleted/re-keyed to 22PWTAdrFgsR) — this sync self-heals that.
 *  - Stores code_id (the real cost-code JobTread id), fixing the T&M bug where
 *    the code NUMBER ("2000"/"8000") was passed to costCodeId instead of the id.
 *  - Flattens the three catalog custom fields (Substrate / Condition / # of Coats)
 *    into columns for fuzzy-match / recommend / create work.
 *
 * IDENTITY: the JobTread cost-item `id` is the true primary key, so we key the
 * upsert on `id`. NAMES ARE NOT UNIQUE — the catalog has Interior/Exterior pairs
 * that share a name under different cost codes (e.g. "French Door Paint - Existing
 * - 1-Coat" exists under 2100 Interior Doors AND 2200 Exterior Doors at different
 * prices). The natural compound identity is name+code, but `id` is what we upsert
 * on. An item deleted+recreated in JT gets a NEW id; the old id simply isn't
 * refreshed by the current run and is removed by the orphan sweep below, so the
 * mirror never serves a dead id to a build.
 *
 * USAGE:
 *   JOBTREAD_GRANT_KEY=... \
 *   SUPABASE_URL=... SUPABASE_SERVICE_ROLE_KEY=... \
 *   node scripts/sync-catalog.mjs [--dry-run]
 *
 * Use the SERVICE ROLE key (server-side only — never ship it to the browser);
 * catalog_items has RLS enabled and writes need to bypass it. --dry-run pulls
 * and reports counts without writing.
 */

import { createClient } from '@supabase/supabase-js';

const ORG_ID = '22PWNY9u7qZd';
const JOBTREAD_API_URL = 'https://api.jobtread.com/pave';
const PAGE_SIZE = 25; // larger pages risk "Request Entity Too Large" with custom fields expanded

// Custom-field ids → column. Verified against the live catalog.
const CF = {
  '22PWTCViXMi8': 'substrate',
  '22PWYi74BtMp': 'condition',
  '22PWjDaWGG46': 'coats',
};

const GRANT_KEY = process.env.JOBTREAD_GRANT_KEY;
const SUPABASE_URL = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL;
const SUPABASE_KEY =
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_KEY;
const DRY_RUN = process.argv.includes('--dry-run');

if (!GRANT_KEY) throw new Error('JOBTREAD_GRANT_KEY is required');
if (!DRY_RUN && (!SUPABASE_URL || !SUPABASE_KEY)) {
  throw new Error('SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are required (omit only with --dry-run)');
}

// One Pave call. grantKey MUST be nested under query.$ (top-level is silently
// ignored — Pave returns {organization:null} with HTTP 200).
async function pave(query) {
  const res = await fetch(JOBTREAD_API_URL, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ query: { $: { grantKey: GRANT_KEY }, ...query } }),
  });
  const text = await res.text();
  let data;
  try { data = JSON.parse(text); }
  catch { throw new Error(`Pave returned non-JSON (HTTP ${res.status}): ${text.slice(0, 300)}`); }
  if (data?.organization == null && data?.error == null) {
    throw new Error(`Pave returned null organization (HTTP ${res.status}) — check grantKey nesting/scope`);
  }
  if (data?.error) throw new Error(`Pave error: ${JSON.stringify(data.error).slice(0, 300)}`);
  return data;
}

function flattenCustomFields(node) {
  const out = { substrate: null, condition: null, coats: null };
  for (const cfv of node?.customFieldValues?.nodes || []) {
    const col = CF[cfv?.customField?.id];
    if (col) out[col] = cfv.value ?? null;
  }
  return out;
}

function toRow(node) {
  const cf = flattenCustomFields(node);
  const name = node.name || '';
  const isTM = name.startsWith('Time & Materials');
  return {
    id: node.id,
    name,
    code: node.costCode?.number ?? null,
    code_id: node.costCode?.id ?? null,
    code_name: node.costCode?.name ?? null,
    cost_type_id: node.costType?.id ?? null,
    cost_type_name: node.costType?.name ?? null,
    unit_id: node.unit?.id ?? null,
    unit_name: node.unit?.name ?? null,
    unit_cost: node.unitCost ?? null,
    unit_price: node.unitPrice ?? null,
    substrate: cf.substrate,
    condition: cf.condition,
    coats: cf.coats,
    kind: isTM ? 'tm' : 'catalog',
    synced_at: new Date().toISOString(),
  };
}

async function pullAll() {
  const rows = [];
  let page = null;
  for (let i = 0; i < 100; i++) { // hard cap; ~491 items / 25 = ~20 pages
    const q = {
      organization: {
        $: { id: ORG_ID },
        costItems: {
          $: {
            size: PAGE_SIZE,
            where: { and: [[['job', 'id'], null]] },
            ...(page ? { page } : {}),
          },
          nextPage: {},
          nodes: {
            id: {}, name: {},
            costCode: { id: {}, number: {}, name: {} },
            costType: { id: {}, name: {} },
            unit: { id: {}, name: {} },
            unitCost: {}, unitPrice: {},
            customFieldValues: { $: { size: 10 }, nodes: { customField: { id: {} }, value: {} } },
          },
        },
      },
    };
    const data = await pave(q);
    const conn = data.organization.costItems;
    for (const node of conn.nodes || []) rows.push(toRow(node));
    page = conn.nextPage;
    if (!page) break;
  }
  return rows;
}

async function main() {
  console.log(`[sync-catalog] pulling JT org catalog (org ${ORG_ID})…`);
  const rows = await pullAll();
  const tm = rows.filter((r) => r.kind === 'tm').length;
  console.log(`[sync-catalog] pulled ${rows.length} items (${tm} T&M, ${rows.length - tm} catalog)`);

  // Sanity: every row must carry an id and a code_id (the two link-critical fields).
  const missingId = rows.filter((r) => !r.id);
  const missingCode = rows.filter((r) => !r.code_id);
  if (missingId.length) console.warn(`[sync-catalog] WARNING: ${missingId.length} rows missing id`);
  if (missingCode.length) console.warn(`[sync-catalog] WARNING: ${missingCode.length} rows missing code_id: ${missingCode.map((r) => r.name).join(', ')}`);

  if (DRY_RUN) {
    console.log('[sync-catalog] --dry-run: no writes. Sample row:', rows[0]);
    return;
  }

  const supabase = createClient(SUPABASE_URL, SUPABASE_KEY, { auth: { persistSession: false } });

  // Single run-stamp shared by every row this sync touches, so orphan cleanup
  // is a simple "anything older than this run" delete (robust for any catalog
  // size — no giant id-list filter that can exceed URL limits). This is also
  // how re-keyed items are handled: the new id upserts a fresh row, the old id
  // isn't refreshed, and the sweep removes it.
  const runStamp = new Date().toISOString();
  for (const r of rows) r.synced_at = runStamp;

  // Upsert on id — the JobTread cost-item id is the true PK. Names are NOT
  // unique (Interior/Exterior pairs share names), so name can never be the key.
  const { error: upErr } = await supabase
    .from('catalog_items')
    .upsert(rows, { onConflict: 'id' });
  if (upErr) throw new Error(`Supabase upsert failed: ${upErr.message}`);

  // Orphan cleanup: any row not refreshed by this run is no longer in JT (or was
  // re-keyed). Removing it keeps the mirror from ever serving a dead id to a build.
  const { error: delErr, count } = await supabase
    .from('catalog_items')
    .delete({ count: 'exact' })
    .lt('synced_at', runStamp);
  if (delErr) console.warn(`[sync-catalog] orphan cleanup skipped: ${delErr.message}`);
  else console.log(`[sync-catalog] removed ${count ?? 0} orphaned rows`);

  console.log(`[sync-catalog] done — ${rows.length} items synced.`);
}

main().catch((err) => { console.error('[sync-catalog] FAILED:', err.message); process.exit(1); });
