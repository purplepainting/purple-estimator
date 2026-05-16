import { useEffect, useState } from 'react';
import { supabase } from './supabaseClient.js';
import JobWalkProposal from './JobWalkProposal.jsx';

const PURPLE = '#2C1654';
const GOLD = '#C8963E';

export default function App() {
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(true);
  const [email, setEmail] = useState('');
  const [status, setStatus] = useState(null);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session);
      setLoading(false);
    });
    const { data: sub } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  async function sendLink(e) {
    e.preventDefault();
    setStatus({ kind: 'pending', msg: 'Sending link…' });
    const { error } = await supabase.auth.signInWithOtp({
      email,
      options: { emailRedirectTo: window.location.origin },
    });
    setStatus(error
      ? { kind: 'error', msg: error.message }
      : { kind: 'ok', msg: 'Check your email for the magic link.' });
  }

  if (loading) {
    return (
      <div style={shellStyle}>
        <div style={{ color: 'white', fontFamily: 'system-ui, sans-serif' }}>Loading…</div>
      </div>
    );
  }

  if (!session) {
    return (
      <div style={shellStyle}>
        <form onSubmit={sendLink} style={cardStyle}>
          <h1 style={{ margin: '0 0 4px', color: PURPLE, fontSize: 22 }}>Purple Estimator</h1>
          <p style={{ margin: '0 0 20px', color: '#666', fontSize: 13 }}>Sign in with a magic link.</p>
          <input
            type="email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            placeholder="you@purplepainting.com"
            required
            style={inputStyle}
          />
          <button type="submit" disabled={status?.kind === 'pending'} style={buttonStyle}>
            Send magic link
          </button>
          {status && (
            <p style={{ marginTop: 12, fontSize: 13, color: status.kind === 'error' ? '#c00' : '#060' }}>
              {status.msg}
            </p>
          )}
        </form>
      </div>
    );
  }

  return (
    <>
      <div style={topBarStyle}>
        <span style={{ fontWeight: 600 }}>Purple Estimator</span>
        <span style={{ flex: 1 }} />
        <span style={{ opacity: 0.85, fontSize: 13 }}>{session.user.email}</span>
        <button onClick={() => supabase.auth.signOut()} style={signOutStyle}>Sign out</button>
      </div>
      <JobWalkProposal />
    </>
  );
}

const shellStyle = { minHeight: '100vh', display: 'grid', placeItems: 'center', background: PURPLE, fontFamily: 'system-ui, sans-serif' };
const cardStyle = { background: 'white', padding: 28, borderRadius: 8, width: 360, boxShadow: '0 8px 32px rgba(0,0,0,0.3)' };
const inputStyle = { width: '100%', padding: 10, border: '1px solid #ddd', borderRadius: 4, fontSize: 14, boxSizing: 'border-box' };
const buttonStyle = { marginTop: 12, width: '100%', padding: 12, background: GOLD, color: 'white', border: 'none', borderRadius: 4, fontSize: 14, fontWeight: 600, cursor: 'pointer' };
const topBarStyle = { display: 'flex', alignItems: 'center', gap: 12, padding: '8px 16px', background: PURPLE, color: 'white', fontFamily: 'system-ui, sans-serif' };
const signOutStyle = { background: 'transparent', color: 'white', border: '1px solid rgba(255,255,255,0.4)', borderRadius: 4, padding: '4px 10px', fontSize: 12, cursor: 'pointer' };
