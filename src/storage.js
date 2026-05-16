import { supabase } from './supabaseClient.js';

const SINGLETON_TABLES = {
  scope_library_v3: 'scope_library',
  catalog_ids_v1: 'catalog_ids',
  tier_multipliers_v1: 'tier_multipliers',
};

const PUBLIC_USER_ID = 'public-user';

async function getSingleton(table) {
  const { data, error } = await supabase.from(table).select('value').eq('id', 1).maybeSingle();
  if (error) throw error;
  return data ? { value: JSON.stringify(data.value) } : null;
}

async function setSingleton(table, jsonString) {
  const parsed = JSON.parse(jsonString);
  const { error } = await supabase
    .from(table)
    .upsert({ id: 1, value: parsed, updated_at: new Date().toISOString() });
  if (error) throw error;
}

async function getKv(key) {
  const { data, error } = await supabase
    .from('kv_store')
    .select('value')
    .eq('user_id', PUBLIC_USER_ID)
    .eq('key', key)
    .maybeSingle();
  if (error) throw error;
  return data ? { value: JSON.stringify(data.value) } : null;
}

async function setKv(key, jsonString) {
  const parsed = JSON.parse(jsonString);
  const { error } = await supabase
    .from('kv_store')
    .upsert({ user_id: PUBLIC_USER_ID, key, value: parsed, updated_at: new Date().toISOString() });
  if (error) throw error;
}

window.storage = {
  async get(key) {
    const table = SINGLETON_TABLES[key];
    return table ? getSingleton(table) : getKv(key);
  },
  async set(key, value) {
    const table = SINGLETON_TABLES[key];
    return table ? setSingleton(table, value) : setKv(key, value);
  },
};
