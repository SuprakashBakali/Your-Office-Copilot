export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  try {
    const { targetUrl, headers, body } = req.body;

    if (!targetUrl) {
      return res.status(400).json({ error: 'targetUrl is required' });
    }

    const fetchOptions = {
      method: 'POST',
      headers: headers || {},
      body: typeof body === 'string' ? body : JSON.stringify(body),
    };

    const response = await fetch(targetUrl, fetchOptions);

    // Forward the status
    res.status(response.status);
    
    // Forward relevant headers, but exclude Transfer-Encoding to let Next/Vercel handle chunks
    response.headers.forEach((value, key) => {
      const lowerKey = key.toLowerCase();
      if (lowerKey !== 'transfer-encoding' && lowerKey !== 'content-encoding') {
        res.setHeader(key, value);
      }
    });

    if (!response.body) {
      return res.end();
    }

    // Stream the response back
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
    res.status(500).json({ error: 'Proxy Request Failed', details: error.message });
  }
}
