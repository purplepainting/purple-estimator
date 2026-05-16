export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  const key = process.env.ANTHROPIC_API_KEY;

  // Vercel parses JSON bodies automatically when content-type matches, but a
  // bare browser fetch() sometimes lands here as a raw string. Handle both.
  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch { body = {}; }
  }
  body = body ?? {};

  console.log(`chat proxy: model=${body.model ?? '<missing>'}, has_key=${!!key}`);

  if (!key) {
    res.status(500).json({ error: 'ANTHROPIC_API_KEY not set on server' });
    return;
  }

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
