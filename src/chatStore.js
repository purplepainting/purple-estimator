import { supabase } from './supabaseClient.js';

async function uid() {
  const { data } = await supabase.auth.getSession();
  return data.session?.user?.id ?? null;
}

export async function listSessions() {
  const u = await uid();
  if (!u) return [];
  const { data, error } = await supabase
    .from('chat_sessions')
    .select('id, title, updated_at, created_at')
    .eq('user_id', u)
    .order('updated_at', { ascending: false });
  if (error) throw error;
  return data;
}

export async function createSession(title = 'New chat') {
  const u = await uid();
  if (!u) throw new Error('Not signed in');
  const { data, error } = await supabase
    .from('chat_sessions').insert({ user_id: u, title }).select().single();
  if (error) throw error;
  return data;
}

export async function renameSession(id, title) {
  const { error } = await supabase
    .from('chat_sessions')
    .update({ title, updated_at: new Date().toISOString() })
    .eq('id', id);
  if (error) throw error;
}

export async function deleteSession(id) {
  const { error } = await supabase.from('chat_sessions').delete().eq('id', id);
  if (error) throw error;
}

export async function loadMessages(sessionId) {
  const { data, error } = await supabase
    .from('chat_messages')
    .select('id, role, content, created_at')
    .eq('session_id', sessionId)
    .order('created_at', { ascending: true });
  if (error) throw error;
  return data;
}

export async function appendMessage(sessionId, role, content) {
  const { data, error } = await supabase
    .from('chat_messages')
    .insert({ session_id: sessionId, role, content })
    .select().single();
  if (error) throw error;
  await supabase
    .from('chat_sessions')
    .update({ updated_at: new Date().toISOString() })
    .eq('id', sessionId);
  return data;
}
