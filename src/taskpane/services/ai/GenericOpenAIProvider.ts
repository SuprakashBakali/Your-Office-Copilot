/**
 * GenericOpenAIProvider — works with any OpenAI-compatible API endpoint.
 * Used for custom models where the user specifies their own base URL + key.
 * Compatible with: NVIDIA NIM, OpenAI, Anthropic (via proxy), Groq,
 *                  OpenRouter, Ollama, Together AI, any custom endpoint.
 */
import { BaseAIProvider, AIRequestOptions, AIResponse, AIStreamChunk, parseOpenAISSEStream } from './types';

export class GenericOpenAIProvider extends BaseAIProvider {
  readonly id = 'generic';
  readonly name = 'Custom (OpenAI-compatible)';
  private _baseUrl: string;

  constructor(baseUrl: string, apiKey = '') {
    super();
    this._baseUrl = baseUrl.replace(/\/$/, ''); // strip trailing slash
    this.apiKey = apiKey;
  }

  get baseUrl(): string { return this._baseUrl; }
  setBaseUrl(url: string) { this._baseUrl = url.replace(/\/$/, ''); }

  requiresKey(): boolean { return true; }
  getModels() { return []; }

  private getHeaders(streaming = false) {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
    };
    if (this.apiKey) {
      headers['Authorization'] = `Bearer ${this.apiKey}`;
    }
    if (streaming) {
      headers['Accept'] = 'text/event-stream';
    }
    return headers;
  }

  async chat(options: AIRequestOptions): Promise<AIResponse> {
    const url = `${this._baseUrl}/chat/completions`;
    const response = await fetch(url, {
      method: 'POST',
      headers: this.getHeaders(false),
      body: JSON.stringify({
        model: options.model,
        messages: options.messages,
        temperature: options.temperature ?? 0.7,
        max_tokens: options.maxTokens ?? 1024,
        stream: false,
      }),
    });

    if (!response.ok) {
      const errText = await response.text().catch(() => '');
      const errSnip = errText.substring(0, 300);
      throw new Error(`API Error ${response.status}: ${errSnip || response.statusText}`);
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
    const url = `${this._baseUrl}/chat/completions`;
    const response = await fetch(url, {
      method: 'POST',
      headers: this.getHeaders(true),
      body: JSON.stringify({
        model: options.model,
        messages: options.messages,
        temperature: options.temperature ?? 0.7,
        max_tokens: options.maxTokens ?? 1024,
        stream: true,
      }),
    });

    if (!response.ok) {
      const errText = await response.text().catch(() => '');
      throw new Error(`API Error ${response.status}: ${errText.substring(0, 300) || response.statusText}`);
    }

    yield* parseOpenAISSEStream(response);
  }
}
