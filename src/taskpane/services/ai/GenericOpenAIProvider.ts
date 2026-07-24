/**
 * GenericOpenAIProvider — works with any OpenAI-compatible API endpoint.
 * Routes all requests through /api/proxy (Vercel serverless) to bypass CORS
 * restrictions in Office Add-in webviews.
 */
import { BaseAIProvider, AIRequestOptions, AIResponse, AIStreamChunk, parseOpenAISSEStream } from './types';

const PROXY_URL = '/api/proxy';

export class GenericOpenAIProvider extends BaseAIProvider {
  readonly id = 'generic';
  readonly name = 'Custom (OpenAI-compatible)';
  private _baseUrl: string;

  constructor(baseUrl: string, apiKey = '') {
    super();
    this._baseUrl = baseUrl.replace(/\/$/, '');
    this.apiKey = apiKey;
  }

  get baseUrl(): string { return this._baseUrl; }
  setBaseUrl(url: string) { this._baseUrl = url.replace(/\/$/, ''); }

  requiresKey(): boolean { return true; }
  getModels() { return []; }

  async chat(options: AIRequestOptions): Promise<AIResponse> {
    const targetUrl = `${this._baseUrl}/chat/completions`;
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${this.apiKey}`,
    };
    if (this._baseUrl.includes('openrouter.ai')) {
      headers['HTTP-Referer'] = 'https://github.com/SuprakashBakali/office-ai-copilot';
      headers['X-Title'] = 'Office AI Copilot';
    }

    const response = await fetch(PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        targetUrl,
        headers,
        body: {
          model: options.model,
          messages: options.messages,
          temperature: options.temperature ?? 0.7,
          top_p: 0.7,
          max_tokens: options.maxTokens ?? 1024,
          stream: false,
        }
      }),
    });

    if (!response.ok) {
      const errText = await response.text().catch(() => '');
      throw new Error(`API Error ${response.status}: ${errText.substring(0, 300) || response.statusText}`);
    }

    const data = await response.json();
    const content = data.choices?.[0]?.message?.content ?? '';
    return {
      content,
      model: data.model ?? options.model,
      tokens: data.usage ? {
        prompt: data.usage.prompt_tokens ?? 0,
        completion: data.usage.completion_tokens ?? 0,
        total: data.usage.total_tokens ?? 0,
      } : undefined,
    };
  }

  async *chatStream(options: AIRequestOptions): AsyncGenerator<AIStreamChunk> {
    const targetUrl = `${this._baseUrl}/chat/completions`;
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${this.apiKey}`,
      'Accept': 'text/event-stream',
    };
    if (this._baseUrl.includes('openrouter.ai')) {
      headers['HTTP-Referer'] = 'https://github.com/SuprakashBakali/office-ai-copilot';
      headers['X-Title'] = 'Office AI Copilot';
    }

    const response = await fetch(PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        targetUrl,
        headers,
        body: {
          model: options.model,
          messages: options.messages,
          temperature: options.temperature ?? 0.7,
          top_p: 0.7,
          max_tokens: options.maxTokens ?? 1024,
          stream: true,
        }
      }),
    });

    if (!response.ok) {
      const errText = await response.text().catch(() => '');
      throw new Error(`API Error ${response.status}: ${errText.substring(0, 300) || response.statusText}`);
    }

    yield* parseOpenAISSEStream(response);
  }
}
