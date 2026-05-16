export default async function handler(req, res) {
  console.log('chat proxy invoked, has_key:', !!process.env.ANTHROPIC_API_KEY, 'method:', req.method);

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

  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch { body = {}; }
  }
  body = body ?? {};
  console.log('chat proxy: model:', body.model, 'max_tokens:', body.max_tokens, 'msg_count:', body.messages?.length);

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

    const rawText = await upstream.text();
    console.log('chat proxy upstream status:', upstream.status, 'body preview:', rawText.slice(0, 800));

    res.status(upstream.status)
       .setHeader('content-type', 'application/json')
       .send(rawText);
  } catch (err) {
    console.error('chat proxy fatal:', err.message, err.stack);
    res.status(500).json({ error: 'proxy_fatal', message: err.message, stack: err.stack?.slice(0, 500) });
  }
}
