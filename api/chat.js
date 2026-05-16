export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }
  const key = process.env.ANTHROPIC_API_KEY;
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
      body: JSON.stringify(req.body),
    });
    const text = await upstream.text();
    res.status(upstream.status)
       .setHeader('content-type', upstream.headers.get('content-type') ?? 'application/json')
       .send(text);
  } catch (e) {
    res.status(502).json({ error: 'Upstream fetch failed', detail: String(e) });
  }
}
