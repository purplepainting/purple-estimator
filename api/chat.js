export default async function handler(req, res) {
  console.log('chat proxy invoked, has_key:', !!process.env.ANTHROPIC_API_KEY, 'len:', (process.env.ANTHROPIC_API_KEY || '').length);

  if (req.method === 'GET' && req.query?.debug === '1') {
    const k = process.env.ANTHROPIC_API_KEY || '';
    res.status(200).json({
      has_anthropic_key: !!k,
      key_length: k.length,
      key_prefix: k.slice(0, 10),
      key_starts_with_sk_ant: k.startsWith('sk-ant-'),
      env_keys_count: Object.keys(process.env).length,
    });
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) {
    res.status(500).json({ error: 'ANTHROPIC_API_KEY not set on server' });
    return;
  }

  // Vercel parses JSON bodies automatically when content-type matches, but a
  // bare browser fetch() sometimes lands here as a raw string. Handle both.
  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch { body = {}; }
  }
  body = body ?? {};

  try {
    const upstream = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': key,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify(body),
    });

    const data = await upstream.json();

    if (!upstream.ok) {
      console.error('chat proxy upstream error:', upstream.status, data);
    }

    res.status(upstream.status).json(data);
  } catch (err) {
    console.error('chat proxy fatal:', err);
    res.status(502).json({ error: err.message });
  }
}
