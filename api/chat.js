const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';
const ANTHROPIC_VERSION = '2023-06-01';

export const config = {
  api: { bodyParser: false },
};

function readRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (chunk) => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf-8')));
    req.on('error', reject);
  });
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();

  if (req.method === 'GET' && req.query?.debug === '1') {
    const k = process.env.ANTHROPIC_API_KEY || '';
    return res.status(200).json({
      has_anthropic_key: !!k,
      key_length: k.length,
      key_prefix: k.slice(0, 10),
      key_starts_with_sk_ant: k.startsWith('sk-ant-'),
      runtime: 'node-esm-rawbody',
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

  // Log what came in so we can SEE what's actually triggering invalid media type
  console.log('chat proxy: method:', req.method, 'content-type:', req.headers['content-type'], 'content-length:', req.headers['content-length']);

  let rawBody;
  try {
    rawBody = await readRawBody(req);
  } catch (err) {
    console.error('readRawBody failed:', err.message);
    return res.status(400).json({ error: 'failed_to_read_body', message: err.message });
  }

  let body;
  try {
    body = JSON.parse(rawBody);
  } catch (err) {
    console.error('JSON.parse failed:', err.message, 'preview:', rawBody.slice(0, 200));
    return res.status(400).json({ error: 'invalid_json', message: err.message, preview: rawBody.slice(0, 200) });
  }

  const { model, max_tokens, system, messages, tools } = body || {};
  if (!model) return res.status(400).json({ error: 'Missing required field: model' });
  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    return res.status(400).json({ error: 'Missing or empty required field: messages' });
  }

  console.log('chat proxy: model:', model, 'max_tokens:', max_tokens, 'msg_count:', messages.length, 'tool_count:', Array.isArray(tools) ? tools.length : 0);

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
        ...(Array.isArray(tools) && tools.length > 0 ? { tools } : {}),
        messages,
      }),
    });
    const data = await upstream.json();
    console.log('chat proxy upstream status:', upstream.status);
    return res.status(upstream.status).json(data);
  } catch (err) {
    console.error('chat proxy fetch error:', err.message);
    return res.status(502).json({ error: 'proxy_fetch_failed', message: err.message });
  }
}
