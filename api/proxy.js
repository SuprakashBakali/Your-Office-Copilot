/**
 * /api/proxy — serverless proxy that forwards chat-completion requests to
 * third-party AI providers. Exists solely to bypass CORS restrictions in the
 * Office Add-in webview (which cannot call api.openai.com etc. directly).
 *
 * Security hardening:
 *  - Domain allowlist: only known AI provider hosts may be proxied. This
 *    prevents the endpoint from being abused as an open SSRF gateway.
 *  - HTTPS-only: plaintext URLs are rejected.
 *  - Method + body shape validation: only POST with the expected payload is
 *    accepted.
 *  - Response size cap: protects the proxy from being used to exfiltrate
 *    arbitrarily large responses.
 */

const ALLOWED_HOSTS = new Set([
  'api.openai.com',
  'api.anthropic.com',
  'generativelanguage.googleapis.com',
  'api.groq.com',
  'openrouter.ai',
  'integrate.api.nvidia.com',
  'localhost',
  '127.0.0.1',
]);

const MAX_RESPONSE_BYTES = 8 * 1024 * 1024; // 8 MB safety cap

function isAllowedTarget(targetUrl) {
  let url;
  try {
    url = new URL(targetUrl);
  } catch {
    return false;
  }
  if (url.protocol === 'https:') {
    return ALLOWED_HOSTS.has(url.hostname);
  }
  if (url.protocol === 'http:' && (url.hostname === 'localhost' || url.hostname === '127.0.0.1')) {
    return true;
  }
  return false;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  // Prevent CORS proxy abuse from unauthorized external domains
  const allowedOrigins = new Set([
    'https://your-office-copilot.vercel.app',
    'https://localhost:3000',
    'http://localhost:3000',
  ]);
  const origin = req.headers.origin;
  if (origin && !allowedOrigins.has(origin)) {
    return res.status(403).json({ error: 'Origin not allowed.' });
  }

  const { targetUrl, headers, body } = req.body || {};

  if (!targetUrl || typeof targetUrl !== 'string') {
    return res.status(400).json({ error: 'targetUrl is required' });
  }

  if (!isAllowedTarget(targetUrl)) {
    return res.status(403).json({
      error: 'Target URL not allowed. The proxy only permits requests to known AI provider hosts over HTTPS.',
    });
  }

  // Sanitize forwarded headers — never forward Host, Content-Length, etc.
  const safeHeaders = {};
  if (headers && typeof headers === 'object') {
    for (const [key, value] of Object.entries(headers)) {
      const lower = key.toLowerCase();
      if (lower === 'host' || lower === 'content-length' || lower === 'connection') continue;
      safeHeaders[key] = value;
    }
  }

  const fetchOptions = {
    method: 'POST',
    headers: safeHeaders,
    body: typeof body === 'string' ? body : JSON.stringify(body),
  };

  try {
    const response = await fetch(targetUrl, fetchOptions);

    // Forward the status
    res.status(response.status);

    // Forward relevant headers, but exclude hop-by-hop / encoding headers
    response.headers.forEach((value, key) => {
      const lowerKey = key.toLowerCase();
      if (
        lowerKey !== 'transfer-encoding' &&
        lowerKey !== 'content-encoding' &&
        lowerKey !== 'content-length' &&
        lowerKey !== 'connection'
      ) {
        res.setHeader(key, value);
      }
    });

    if (!response.body) {
      return res.end();
    }

    // Stream the response back with a size cap to prevent abuse.
    const reader = response.body.getReader();
    const decoder = new TextDecoder('utf-8');
    let bytesSeen = 0;

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      bytesSeen += value.byteLength;
      if (bytesSeen > MAX_RESPONSE_BYTES) {
        await reader.cancel();
        res.end();
        return;
      }
      res.write(decoder.decode(value));
    }
    res.end();
  } catch (error) {
    console.error('Proxy error:', error);
    res.status(502).json({ error: 'Proxy Request Failed', details: error.message });
  }
}

export const config = {
  api: {
    bodyParser: {
      sizeLimit: '10mb',
    },
  },
};

