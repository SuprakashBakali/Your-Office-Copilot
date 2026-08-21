/**
 * OpenAICompatibleProvider — shared base for any provider that speaks the
 * OpenAI Chat Completions API (`POST {baseUrl}/chat/completions`, SSE stream
 * format `data: {...}\\n\\n`, terminated by `data: [DONE]`).
 *
 * Concrete providers only need to declare `id`, `name`, `baseUrl`, the model
 * list, and whether a key is required. They may optionally override
 * `buildHeaders()` / `buildBody()` to add provider-specific quirks.
 */
import { BaseAIProvider, AIRequestOptions, AIResponse, AIStreamChunk, parseOpenAISSEStream } from './types';

const PROXY_URL = '/api/proxy';

/** Models that use max_completion_tokens instead of max_tokens and don't
 *  support temperature (OpenAI o-series reasoning models). */
const REASONING_MODEL_PREFIXES = ['o1', 'o3', 'o4'];
function isReasoningModel(modelId: string): boolean {
  return REASONING_MODEL_PREFIXES.some(p =>
    modelId === p || modelId.startsWith(`${p}-`) || modelId.startsWith(`${p}/`)
  );
}

export abstract class OpenAICompatibleProvider extends BaseAIProvider {
  /** When true (default), requests are tunnelled through /api/proxy to bypass
   *  CORS in the Office Add-in webview. Local providers (e.g. Ollama) can set
   *  this to false to hit the endpoint directly. */
  protected useProxy: boolean = true;

  /** Set to false for providers that reject stream_options (e.g. Groq). */
  protected supportsStreamOptions: boolean = true;

  /** Override to inject provider-specific headers (e.g. OpenRouter's
   *  HTTP-Referer / X-Title). */
  protected buildHeaders(_options: AIRequestOptions): Record<string, string> {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
    };
    if (this.requiresKey() && this.apiKey) {
      headers['Authorization'] = `Bearer ${this.apiKey}`;
    }
    return headers;
  }

  /** Override to mutate the JSON body (e.g. add `plugins` for OpenRouter web
   *  search, drop `top_p` for providers that reject it). */
  protected buildBody(options: AIRequestOptions): Record<string, unknown> {
    const reasoning = isReasoningModel(options.model);
    const body: Record<string, unknown> = {
      model: options.model,
      messages: options.messages,
    };

    if (reasoning) {
      // o1/o3/o4 use max_completion_tokens and don't accept temperature
      if (options.maxTokens) body['max_completion_tokens'] = options.maxTokens;
    } else {
      body['temperature'] = options.temperature ?? 0.7;
      body['max_tokens'] = options.maxTokens ?? 2048;
    }

    return body;
  }

  private buildRequest(options: AIRequestOptions, stream: boolean) {
    const headers = this.buildHeaders(options);
    const bodyObj = { ...this.buildBody(options), stream };

    // For streaming, request JSON event stream and include usage data if possible
    if (stream) {
      headers['Accept'] = 'text/event-stream';
      // stream_options supported by OpenAI and most compatible providers,
      // but rejected by Groq and some others.
      if (this.supportsStreamOptions) {
        (bodyObj as any)['stream_options'] = { include_usage: true };
      }
    }

    const targetUrl = `${this.baseUrl}/chat/completions`;

    const init: RequestInit = {
      method: 'POST',
      signal: options.signal,
    };

    if (this.useProxy) {
      init.headers = { 'Content-Type': 'application/json' };
      init.body = JSON.stringify({ targetUrl, headers, body: bodyObj });
      return { url: PROXY_URL, init };
    }
    init.headers = headers;
    init.body = JSON.stringify(bodyObj);
    return { url: targetUrl, init };
  }

  protected async throwIfError(response: Response): Promise<void> {
    if (response.ok) return;
    let errMsg = `${response.status} ${response.statusText}`;
    try {
      const raw = await response.text();
      if (raw) {
        // Try to parse structured error (OpenAI, Groq, NVIDIA, etc.)
        try {
          const parsed = JSON.parse(raw);
          const msg = parsed?.error?.message || parsed?.message || parsed?.detail;
          if (msg) {
            errMsg = `${response.status} — ${msg}`;
          } else {
            errMsg = `${response.status} — ${raw.substring(0, 300)}`;
          }
        } catch {
          errMsg = `${response.status} — ${raw.substring(0, 300)}`;
        }
      }
    } catch { /* ignore body read error */ }
    throw new Error(`${this.name} API Error: ${errMsg}`);
  }

  async chat(options: AIRequestOptions): Promise<AIResponse> {
    const { url, init } = this.buildRequest(options, false);
    const response = await fetch(url, init);
    await this.throwIfError(response);

    const data = await response.json();
    const choice = data.choices?.[0];
    const content = choice?.message?.content ?? '';
    // DeepSeek-R1 and similar models return reasoning in a separate field
    const thinking = choice?.message?.reasoning_content ?? undefined;

    return {
      content,
      thinking,
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
