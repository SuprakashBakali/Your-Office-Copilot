/**
 * OpenAICompatibleProvider — shared base for any provider that speaks the
 * OpenAI Chat Completions API (`POST {baseUrl}/chat/completions`, SSE stream
 * format `data: {...}\n\n`, terminated by `data: [DONE]`).
 *
 * Concrete providers only need to declare `id`, `name`, `baseUrl`, the model
 * list, and whether a key is required. They may optionally override
 * `buildHeaders()` / `buildBody()` to add provider-specific quirks.
 *
 * This eliminates ~500 lines of duplicated fetch + SSE parsing code across
 * NVIDIA, OpenAI, Groq, OpenRouter, and Ollama providers.
 */
import { BaseAIProvider, AIRequestOptions, AIResponse, AIStreamChunk, parseOpenAISSEStream } from './types';

const PROXY_URL = '/api/proxy';

export abstract class OpenAICompatibleProvider extends BaseAIProvider {
  /** When true (default), requests are tunneled through /api/proxy to bypass
   *  CORS in the Office Add-in webview. Local providers (e.g. Ollama) can set
   *  this to false to hit the endpoint directly. */
  protected useProxy: boolean = true;

  /** Override to inject provider-specific headers (e.g. OpenRouter's
   *  HTTP-Referer / X-Title). */
  protected buildHeaders(_options: AIRequestOptions): Record<string, string> {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
    };
    if (this.requiresKey()) {
      headers['Authorization'] = `Bearer ${this.apiKey}`;
    }
    return headers;
  }

  /** Override to mutate the JSON body (e.g. add `plugins` for OpenRouter web
   *  search, drop `top_p` for providers that reject it). */
  protected buildBody(options: AIRequestOptions): Record<string, unknown> {
    const body: Record<string, unknown> = {
      model: options.model,
      messages: options.messages,
      temperature: options.temperature ?? 0.7,
      max_tokens: options.maxTokens ?? 1024,
    };
    return body;
  }

  private buildRequest(options: AIRequestOptions, stream: boolean) {
    const headers = this.buildHeaders(options);
    if (stream) headers['Accept'] = 'text/event-stream';
    const body = { ...this.buildBody(options), stream };
    const targetUrl = `${this.baseUrl}/chat/completions`;

    const init: RequestInit = {
      method: 'POST',
      // Thread the AbortSignal through so `cancelStream()` actually cancels
      // the in-flight fetch instead of just stopping the consumer loop.
      signal: options.signal,
    };

    if (this.useProxy) {
      init.headers = { 'Content-Type': 'application/json' };
      init.body = JSON.stringify({ targetUrl, headers, body });
      return { url: PROXY_URL, init };
    }
    init.headers = headers;
    init.body = JSON.stringify(body);
    return { url: targetUrl, init };
  }

  private async throwIfError(response: Response): Promise<void> {
    if (response.ok) return;
    const errBody = await response.text().catch(() => '');
    const errMsg = errBody
      ? `${response.status} — ${errBody.substring(0, 200)}`
      : `${response.status} ${response.statusText}`;
    throw new Error(`${this.name} API Error: ${errMsg}`);
  }

  async chat(options: AIRequestOptions): Promise<AIResponse> {
    const { url, init } = this.buildRequest(options, false);
    const response = await fetch(url, init);
    await this.throwIfError(response);

    const data = await response.json();
    return {
      content: data.choices?.[0]?.message?.content ?? '',
      model: data.model ?? options.model,
      tokens: data.usage
        ? {
            prompt: data.usage.prompt_tokens ?? 0,
            completion: data.usage.completion_tokens ?? 0,
            total: data.usage.total_tokens ?? 0,
          }
        : undefined,
    };
  }

  async *chatStream(options: AIRequestOptions): AsyncGenerator<AIStreamChunk> {
    const { url, init } = this.buildRequest(options, true);
    const response = await fetch(url, init);
    await this.throwIfError(response);
    yield* parseOpenAISSEStream(response, options.signal);
  }
}
