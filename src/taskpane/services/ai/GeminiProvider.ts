import { BaseAIProvider, AIRequestOptions, AIResponse, AIStreamChunk } from './types';

export class GeminiProvider extends BaseAIProvider {
  readonly id = 'gemini';
  readonly name = 'Google Gemini';
  readonly baseUrl = 'https://generativelanguage.googleapis.com/v1beta';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      // Updated to real Gemini API model IDs (2025).
      // Previous IDs gemini-3.6-flash / gemini-3.1-pro-preview do NOT exist
      // on the API and caused immediate 404 errors.
      { id: 'gemini-2.5-flash', name: 'Gemini 2.5 Flash' },
      { id: 'gemini-2.5-pro', name: 'Gemini 2.5 Pro' },
      { id: 'gemini-2.0-flash', name: 'Gemini 2.0 Flash' },
      { id: 'gemini-1.5-flash', name: 'Gemini 1.5 Flash' },
      { id: 'gemini-1.5-pro', name: 'Gemini 1.5 Pro' },
    ];
  }

  private formatRequest(options: AIRequestOptions) {
    const contents = [];
    let systemInstruction = null;

    for (const msg of options.messages) {
      if (msg.role === 'system') {
        systemInstruction = { parts: [{ text: msg.content }] };
      } else {
        contents.push({
          role: msg.role === 'user' ? 'user' : 'model',
          parts: [{ text: msg.content }]
        });
      }
    }

    const payload: any = { contents };
    if (systemInstruction) {
      payload.systemInstruction = systemInstruction;
    }

    payload.generationConfig = {
      temperature: options.temperature,
      maxOutputTokens: options.maxTokens
    };

    // Enable Google Search Grounding if webSearch is true
    if (options.webSearch) {
      payload.tools = [{ googleSearch: {} }];
    }

    return payload;
  }

  /** Build headers with the API key in the x-goog-api-key header instead of
   *  the URL query string. Putting the key in the URL exposes it in DevTools
   *  network tab, Referer headers, and server access logs. */
  private getHeaders() {
    return {
      'Content-Type': 'application/json',
      'x-goog-api-key': this.apiKey,
    };
  }

  async chat(options: AIRequestOptions): Promise<AIResponse> {
    const payload = this.formatRequest(options);
    // Key is now in the header, NOT in the URL.
    const url = `${this.baseUrl}/models/${options.model}:generateContent`;

    const response = await fetch(url, {
      method: 'POST',
      headers: this.getHeaders(),
      signal: options.signal,
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errBody = await response.text().catch(() => '');
      const errMsg = errBody ? `${response.status} — ${errBody.substring(0, 200)}` : `${response.status} ${response.statusText}`;
      throw new Error(`Gemini API Error: ${errMsg}`);
    }

    const data = await response.json();
    return {
      content: data.candidates?.[0]?.content?.parts?.[0]?.text || '',
      model: options.model,
      tokens: {
        prompt: data.usageMetadata?.promptTokenCount ?? 0,
        completion: data.usageMetadata?.candidatesTokenCount ?? 0,
        total: data.usageMetadata?.totalTokenCount ?? 0
      }
    };
  }

  async *chatStream(options: AIRequestOptions): AsyncGenerator<AIStreamChunk> {
    const payload = this.formatRequest(options);
    // Key is now in the header, NOT in the URL.
    const url = `${this.baseUrl}/models/${options.model}:streamGenerateContent?alt=sse`;

    const response = await fetch(url, {
      method: 'POST',
      headers: this.getHeaders(),
      signal: options.signal,
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errBody = await response.text().catch(() => '');
      const errMsg = errBody ? `${response.status} — ${errBody.substring(0, 200)}` : `${response.status} ${response.statusText}`;
      throw new Error(`Gemini API Error: ${errMsg}`);
    }

    if (!response.body) throw new Error("No response body");

    const reader = response.body.getReader();
    const onAbort = () => { reader.cancel().catch(() => {}); };
    if (options.signal) {
      if (options.signal.aborted) { await reader.cancel().catch(() => {}); return; }
      options.signal.addEventListener('abort', onAbort, { once: true });
    }
    const decoder = new TextDecoder("utf-8");
    let buffer = "";

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || "";

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed || !trimmed.startsWith("data: ")) continue;

          const dataStr = trimmed.slice(6);
          if (dataStr === "[DONE]") {
            yield { content: "", done: true };
            return;
          }

          try {
            const data = JSON.parse(dataStr);
            const content = data.candidates?.[0]?.content?.parts?.[0]?.text;
            if (content) {
              yield { content, done: false };
            }
          } catch (e) {
            console.error("Failed to parse Gemini SSE chunk", e);
          }
        }
      }
      yield { content: "", done: true };
    } finally {
      if (options.signal) options.signal.removeEventListener('abort', onAbort);
      reader.releaseLock();
    }
  }
}
