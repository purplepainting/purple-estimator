const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';
const ANTHROPIC_VERSION = '2023-06-01';

export default async function handler(req, res) {
  // CORS / preflight
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  // Debug endpoint
  if (req.method === 'GET' && req.query?.debug === '1') {
    const k = process.env.ANTHROPIC_API_KEY || '';
    return res.status(200).json({
      has_anthropic_key: !!k,
      key_length: k.length,
      key_prefix: k.slice(0, 10),
      key_starts_with_sk_ant: k.startsWith('sk-ant-'),
      runtime: 'node-esm',
      node_version: process.version,
    });
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed. Use POST.' });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY not set on server' });
  }

  // Vercel auto-parses JSON bodies for application/json
  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch { body = {}; }
  }
  body = body || {};

  const { model, max_tokens, system, messages } = body;

  if (!model) {
    return res.status(400).json({ error: 'Missing required field: model' });
  }
  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    return res.status(400).json({ error: 'Missing or empty required field: messages' });
  }

  console.log('chat proxy: model:', model, 'max_tokens:', max_tokens, 'msg_count:', messages.length);

  try {
    const upstream = await fetch(ANTHROPIC_API_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': ANTHROPIC_VERSION,
      },
      body: JSON.stringify({
        model,
        max_tokens: max_tokens ?? 4096,
        ...(system ? { system } : {}),
        messages,
      }),
    });

    const data = await upstream.json();
    console.log('chat proxy upstream status:', upstream.status);
    return res.status(upstream.status).json(data);
  } catch (err) {
    console.error('chat proxy fetch error:', err.message, err.stack);
    return res.status(502).json({
      error: 'proxy_fetch_failed',
      message: err.message,
    });
  }
}
