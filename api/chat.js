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

  const { model, max_tokens, system, messages, tools, stream } = body || {};
  if (!model) return res.status(400).json({ error: 'Missing required field: model' });
  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    return res.status(400).json({ error: 'Missing or empty required field: messages' });
  }

  const isStream = stream === true;
  console.log('chat proxy: model:', model, 'max_tokens:', max_tokens, 'msg_count:', messages.length, 'tool_count:', Array.isArray(tools) ? tools.length : 0, 'stream:', isStream);

  // Prompt caching: mark the system prompt and the tools block as ephemeral so
  // Anthropic caches them across the build-loop turns. Cached input tokens
  // don't count toward ITPM and bill at ~10% of base. Caching applies to the
  // request prefix (system → tools → messages); messages stay uncached because
  // they change every turn. Note: caching only engages when the cached prefix
  // is ≥ ~1024 tokens; below that it silently no-ops.
  let cachedSystem;
  if (typeof system === 'string' && system.length > 0) {
    cachedSystem = [{ type: 'text', text: system, cache_control: { type: 'ephemeral' } }];
  } else if (Array.isArray(system) && system.length > 0) {
    // Already a content-block array — pass through unchanged.
    cachedSystem = system;
  }
  const cachedTools = Array.isArray(tools) && tools.length > 0
    ? tools.map((t, i) =>
        i === tools.length - 1 ? { ...t, cache_control: { type: 'ephemeral' } } : t
      )
    : null;

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
        ...(cachedSystem ? { system: cachedSystem } : {}),
        ...(cachedTools ? { tools: cachedTools } : {}),
        ...(isStream ? { stream: true } : {}),
        messages,
      }),
    });

    // Streaming pass-through (opt-in via body.stream === true). Bytes flow as
    // they arrive, so the function never goes idle long enough for Vercel to
    // kill it with FUNCTION_INVOCATION_TIMEOUT. The buffered path below stays
    // BYTE-FOR-BYTE unchanged for non-streaming callers (BuildChat /
    // DocumentChat / categorize / ClarifyChat) — they don't set stream and
    // still get a single JSON response.
    if (isStream) {
      console.log('chat proxy upstream status (stream):', upstream.status);
      if (!upstream.ok) {
        // Upstream rejected the stream request (auth, schema, etc.). Anthropic
        // returns a plain JSON error here — pass it through as JSON so the
        // client's existing !response.ok branch picks it up.
        const errText = await upstream.text();
        console.error('chat proxy upstream error (stream req):', upstream.status, errText.slice(0, 200));
        res.status(upstream.status);
        res.setHeader('Content-Type', 'application/json');
        res.send(errText);
        return;
      }
      res.status(200);
      res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
      res.setHeader('Cache-Control', 'no-cache, no-transform');
      res.setHeader('Connection', 'keep-alive');
      const reader = upstream.body.getReader();
      const decoder = new TextDecoder();
      try {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          res.write(decoder.decode(value, { stream: true }));
        }
      } finally {
        res.end();
      }
      return;
    }

    const data = await upstream.json();
    console.log('chat proxy upstream status:', upstream.status);
    console.log('chat proxy cache:', JSON.stringify(data?.usage || {}));
    return res.status(upstream.status).json(data);
  } catch (err) {
    console.error('chat proxy fetch error:', err.message);
    return res.status(502).json({ error: 'proxy_fetch_failed', message: err.message });
  }
}
