import { BaseAIProvider, AIRequestOptions, AIResponse, AIStreamChunk } from './types';

const PROXY_URL = '/api/proxy';

/** Gemini 2.5+ models support thinking — enable thinkingConfig for them. */
function supportsThinking(modelId: string): boolean {
  return modelId.startsWith('gemini-2.5') || modelId.startsWith('gemini-exp');
}

export class GeminiProvider extends BaseAIProvider {
  readonly id = 'gemini';
  readonly name = 'Google Gemini';
  readonly baseUrl = 'https://generativelanguage.googleapis.com/v1beta';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      { id: 'gemini-2.5-flash',           name: 'Gemini 2.5 Flash' },
      { id: 'gemini-2.5-pro',             name: 'Gemini 2.5 Pro' },
      { id: 'gemini-2.5-flash-thinking',  name: 'Gemini 2.5 Flash Thinking' },
      { id: 'gemini-2.0-flash',           name: 'Gemini 2.0 Flash' },
      { id: 'gemini-2.0-flash-lite',      name: 'Gemini 2.0 Flash Lite' },
      { id: 'gemini-1.5-flash',           name: 'Gemini 1.5 Flash' },
      { id: 'gemini-1.5-flash-8b',        name: 'Gemini 1.5 Flash 8B' },
      { id: 'gemini-1.5-pro',             name: 'Gemini 1.5 Pro' },
    ];
  }

  private getHeaders(): Record<string, string> {
    return {
      'Content-Type': 'application/json',
      'x-goog-api-key': this.apiKey,
    };
  }

  private formatRequest(options: AIRequestOptions): Record<string, unknown> {
    const contents: unknown[] = [];
    let systemInstruction: unknown = null;

    for (const msg of options.messages) {
      if (msg.role === 'system') {
        systemInstruction = { parts: [{ text: msg.content }] };
      } else {
        contents.push({
          role: msg.role === 'user' ? 'user' : 'model',
          parts: [{ text: msg.content }],
        });
      }
    }

    const payload: Record<string, unknown> = { contents };
    if (systemInstruction) payload['systemInstruction'] = systemInstruction;

    // generationConfig — only include defined values to avoid 400 errors
    const genConfig: Record<string, unknown> = {};
    if (options.temperature !== undefined) genConfig['temperature'] = options.temperature;
    if (options.maxTokens) genConfig['maxOutputTokens'] = options.maxTokens;
    if (Object.keys(genConfig).length > 0) payload['generationConfig'] = genConfig;

    // Enable Google Search Grounding for web search
    if (options.webSearch) {
      payload['tools'] = [{ googleSearch: {} }];
    }

    // Thinking config for Gemini 2.5+ (budget_tokens = 0 means dynamic/auto)
    if (supportsThinking(options.model)) {
      payload['thinkingConfig'] = { thinkingBudget: 4096 };
    }

    return payload;
  }

  private async throwIfError(response: Response): Promise<void> {
    if (response.ok) return;
    let errMsg = `${response.status} ${response.statusText}`;
    try {
      const raw = await response.text();
      if (raw) {
        try {
          const parsed = JSON.parse(raw);
          const msg = parsed?.error?.message;
          errMsg = msg ? `${response.status} — ${msg}` : `${response.status} — ${raw.substring(0, 300)}`;
        } catch {
          errMsg = `${response.status} — ${raw.substring(0, 300)}`;
        }
      }
    } catch { /* ignore */ }
    throw new Error(`Gemini API Error: ${errMsg}`);
  }

  async chat(options: AIRequestOptions): Promise<AIResponse> {
    const payload = this.formatRequest(options);
    // Route through proxy to avoid CORS — same as Anthropic.
    // Gemini's API at generativelanguage.googleapis.com does NOT include CORS
    // headers for browser webview origins used by Office Add-ins.
    const targetUrl = `${this.baseUrl}/models/${options.model}:generateContent`;

    const response = await fetch(PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal: options.signal,
      body: JSON.stringify({ targetUrl, headers: this.getHeaders(), body: payload }),
    });
    await this.throwIfError(response);

    const data = await response.json();
    // Extract text and thinking from parts array
    let content = '';
    let thinking = '';
    const parts: any[] = data.candidates?.[0]?.content?.parts ?? [];
    for (const part of parts) {
      if (part.thought === true || part.thought) {
        thinking += part.text ?? '';
      } else {
        content += part.text ?? '';
      }
    }
    return {
      content,
      thinking: thinking || undefined,
      model: options.model,
      tokens: {
        prompt: data.usageMetadata?.promptTokenCount ?? 0,
        completion: data.usageMetadata?.candidatesTokenCount ?? 0,
        total: data.usageMetadata?.totalTokenCount ?? 0,
      },
    };
  }

  async *chatStream(options: AIRequestOptions): AsyncGenerator<AIStreamChunk> {
    const payload = this.formatRequest(options);
    const targetUrl = `${this.baseUrl}/models/${options.model}:streamGenerateContent?alt=sse`;

    // Route through proxy for CORS
    const response = await fetch(PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal: options.signal,
      body: JSON.stringify({
        targetUrl,
        headers: { ...this.getHeaders(), 'Accept': 'text/event-stream' },
        body: payload,
      }),
    });
    await this.throwIfError(response);

    if (!response.body) throw new Error('No response body');

    const reader = response.body.getReader();
    const onAbort = () => { reader.cancel().catch(() => {}); };
    if (options.signal) {
      if (options.signal.aborted) { await reader.cancel().catch(() => {}); return; }
      options.signal.addEventListener('abort', onAbort, { once: true });
    }
    const decoder = new TextDecoder('utf-8');
    let buffer = '';

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed || !trimmed.startsWith('data: ')) continue;

          const dataStr = trimmed.slice(6);
          if (dataStr === '[DONE]') {
            yield { content: '', done: true };
            return;
          }

          try {
            const data = JSON.parse(dataStr);
            const parts: any[] = data.candidates?.[0]?.content?.parts ?? [];
            for (const part of parts) {
              if (!part.text) continue;
              // Gemini 2.5 thinking parts have thought=true
              if (part.thought === true || part.thought) {
                yield { content: '', done: false, thinking: part.text };
              } else {
                yield { content: part.text, done: false };
              }
            }
            // Gemini signals end via finishReason
            const finishReason = data.candidates?.[0]?.finishReason;
            if (finishReason && finishReason !== 'STOP' && finishReason !== '') {
              // Non-STOP reasons (e.g. MAX_TOKENS) — still emit done
              if (finishReason === 'MAX_TOKENS' || finishReason === 'RECITATION' || finishReason === 'SAFETY') {
                yield { content: '', done: true };
                return;
              }
            }
          } catch {
            // Ignore malformed chunks
          }
        }
      }
      yield { content: '', done: true };
    } finally {
      if (options.signal) options.signal.removeEventListener('abort', onAbort);
      reader.releaseLock();
    }
  }
}
