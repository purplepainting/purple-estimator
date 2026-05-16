export const config = {
  maxDuration: 60,
};

export default async function handler(request) {
  const url = new URL(request.url);

  if (request.method === 'GET' && url.searchParams.get('debug') === '1') {
    const k = process.env.ANTHROPIC_API_KEY || '';
    return new Response(JSON.stringify({
      has_anthropic_key: !!k,
      key_length: k.length,
      key_prefix: k.slice(0, 10),
      key_starts_with_sk_ant: k.startsWith('sk-ant-'),
      env_keys_count: Object.keys(process.env).length,
      runtime: 'web-api',
    }), { status: 200, headers: { 'content-type': 'application/json' } });
  }

  if (request.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'Method not allowed' }), {
      status: 405, headers: { 'content-type': 'application/json' },
    });
  }

  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) {
    return new Response(JSON.stringify({ error: 'ANTHROPIC_API_KEY not set on server' }), {
      status: 500, headers: { 'content-type': 'application/json' },
    });
  }

  let bodyText;
  try {
    bodyText = await request.text();
  } catch (err) {
    console.error('failed to read request body:', err.message);
    return new Response(JSON.stringify({ error: 'failed_to_read_body', message: err.message }), {
      status: 400, headers: { 'content-type': 'application/json' },
    });
  }

  let body;
  try {
    body = JSON.parse(bodyText);
  } catch (err) {
    console.error('invalid JSON body:', err.message, 'preview:', bodyText.slice(0, 200));
    return new Response(JSON.stringify({ error: 'invalid_json', message: err.message }), {
      status: 400, headers: { 'content-type': 'application/json' },
    });
  }

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
    console.log('chat proxy upstream status:', upstream.status, 'body preview:', rawText.slice(0, 400));

    return new Response(rawText, {
      status: upstream.status,
      headers: { 'content-type': 'application/json' },
    });
  } catch (err) {
    console.error('chat proxy fatal:', err.message, err.stack);
    return new Response(JSON.stringify({
      error: 'proxy_fatal',
      message: err.message,
      stack: err.stack?.slice(0, 500),
    }), { status: 500, headers: { 'content-type': 'application/json' } });
  }
}
