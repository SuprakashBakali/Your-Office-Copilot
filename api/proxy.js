/**
 * Vercel Serverless Proxy — routes AI API requests server-side to avoid CORS issues
 * in Office Add-in webviews.
 *
 * POST /api/proxy
 * Body: { targetUrl, ...openAIChatBody }
 * Headers: Authorization: Bearer <apiKey>
 */

const ALLOWED_DOMAINS = [
  'integrate.api.nvidia.com',
  'api.openai.com',
  'api.anthropic.com',
  'generativelanguage.googleapis.com',
  'api.groq.com',
  'openrouter.ai',
  'localhost',
  '127.0.0.1',
];

export default async function handler(req, res) {
  // CORS headers — allow the Office add-in origin
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, Accept');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { targetUrl, ...requestBody } = req.body || {};

  if (!targetUrl) {
    return res.status(400).json({ error: 'Missing targetUrl' });
  }

  // Validate domain whitelist
  let parsedUrl;
  try {
    parsedUrl = new URL(targetUrl);
  } catch {
    return res.status(400).json({ error: 'Invalid targetUrl' });
  }

  const isAllowed = ALLOWED_DOMAINS.some(d => parsedUrl.hostname === d || parsedUrl.hostname.endsWith('.' + d));
  if (!isAllowed) {
    return res.status(403).json({ error: `Domain not allowed: ${parsedUrl.hostname}` });
  }

  const isStreaming = requestBody.stream === true;

  try {
    const upstreamRes = await fetch(targetUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': req.headers['authorization'] || '',
        ...(isStreaming ? { 'Accept': 'text/event-stream' } : {}),
      },
      body: JSON.stringify(requestBody),
    });

    if (isStreaming) {
      res.setHeader('Content-Type', 'text/event-stream');
      res.setHeader('Cache-Control', 'no-cache');
      res.setHeader('Connection', 'keep-alive');

      const reader = upstreamRes.body.getReader();
      while (true) {
        const { done, value } = await reader.read();
        if (done) { res.end(); break; }
        res.write(Buffer.from(value));
      }
    } else {
      const text = await upstreamRes.text();
      res.status(upstreamRes.status)
        .setHeader('Content-Type', upstreamRes.headers.get('content-type') || 'application/json')
        .send(text);
    }
  } catch (err) {
    res.status(502).json({ error: `Proxy error: ${(err).message}` });
  }
}
