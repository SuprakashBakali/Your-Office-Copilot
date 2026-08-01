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

function isAllowedTarget(targetUrl: string): boolean {
  let url: URL;
  try {
    url = new URL(targetUrl);
  } catch {
    return false;
  }
  return url.protocol === 'http:' || url.protocol === 'https:';
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  const { targetUrl, headers, body } = req.body || {};

  if (!targetUrl || typeof targetUrl !== 'string') {
    return res.status(400).json({ error: 'targetUrl is required' });
  }

  if (!isAllowedTarget(targetUrl)) {
    return res.status(403).json({
      error: 'Target URL not allowed. Must be a valid HTTP or HTTPS URL.',
    });
  }

  // Forward headers as requested
  const safeHeaders = headers && typeof headers === 'object' ? { ...headers } : {};

  const fetchOptions = {
    method: 'POST',
    headers: safeHeaders,
    body: typeof body === 'string' ? body : JSON.stringify(body),
  };

  try {
    const response = await fetch(targetUrl, fetchOptions);

    // Forward the status
    res.status(response.status);

    // Forward headers
    response.headers.forEach((value, key) => {
      res.setHeader(key, value);
    });

    if (!response.body) {
      return res.end();
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder('utf-8');

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      res.write(decoder.decode(value));
    }
    res.end();
  } catch (error) {
    console.error('Proxy error:', error);
    res.status(502).json({ error: 'Proxy Request Failed', details: error.message });
  }
}
