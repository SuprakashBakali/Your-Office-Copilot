/**
 * GenericOpenAIProvider — works with any user-supplied OpenAI-compatible API
 * endpoint. Used by the "My Models" settings panel for custom provider
 * configurations (custom baseUrl + key + modelId).
 *
 * Smart endpoint detection:
 * - If the baseUrl points to api.anthropic.com → delegates to AnthropicProvider
 * - If the baseUrl points to generativelanguage.googleapis.com → delegates to GeminiProvider
 * - Otherwise → standard OpenAI-compatible /chat/completions endpoint
 */
import { OpenAICompatibleProvider } from './OpenAICompatibleProvider';
import { AIRequestOptions, AIResponse, AIStreamChunk } from './types';

export class GenericOpenAIProvider extends OpenAICompatibleProvider {
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

  /** Detect whether this endpoint is the native Anthropic API. */
  private isAnthropic(): boolean {
    return this._baseUrl.includes('api.anthropic.com');
  }

  /** Detect whether this endpoint is the native Google Gemini API. */
  private isGemini(): boolean {
    return this._baseUrl.includes('generativelanguage.googleapis.com');
  }

  /** Detect whether this endpoint is the native NVIDIA NIM API. */
  private isNvidia(): boolean {
    return this._baseUrl.includes('integrate.api.nvidia.com');
  }

  /** Lazily instantiate the native Anthropic provider with our key. */
  private async getAnthropicProvider() {
    const { AnthropicProvider } = await import('./AnthropicProvider');
    const p = new AnthropicProvider();
    p.setApiKey(this.apiKey);
    return p;
  }

  /** Lazily instantiate the native Gemini provider with our key. */
  private async getGeminiProvider() {
    const { GeminiProvider } = await import('./GeminiProvider');
    const p = new GeminiProvider();
    p.setApiKey(this.apiKey);
    return p;
  }

  /** Lazily instantiate the native NVIDIA provider with our key. */
  private async getNvidiaProvider() {
    const { NvidiaProvider } = await import('./NvidiaProvider');
    const p = new NvidiaProvider();
    p.setApiKey(this.apiKey);
    return p;
  }

  /** OpenRouter attribution headers. */
  protected buildHeaders(_options: AIRequestOptions): Record<string, string> {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${this.apiKey}`,
    };
    if (this._baseUrl.includes('openrouter.ai')) {
      headers['HTTP-Referer'] = 'https://github.com/SuprakashBakali/office-ai-copilot';
      headers['X-Title'] = 'Office AI Copilot';
    }
    return headers;
  }

  /** OpenRouter web-search plugin support. */
  protected buildBody(options: AIRequestOptions): Record<string, unknown> {
    const body = super.buildBody(options);
    if (options.webSearch && this._baseUrl.includes('openrouter.ai')) {
      body['plugins'] = [{ id: 'web', max_results: 5 }];
    }
    return body;
  }

  async chat(options: AIRequestOptions): Promise<AIResponse> {
    if (this.isAnthropic()) {
      const p = await this.getAnthropicProvider();
      return p.chat(options);
    }
    if (this.isGemini()) {
      const p = await this.getGeminiProvider();
      return p.chat(options);
    }
    if (this.isNvidia()) {
      const p = await this.getNvidiaProvider();
      return p.chat(options);
    }
    return super.chat(options);
  }

  async *chatStream(options: AIRequestOptions): AsyncGenerator<AIStreamChunk> {
    if (this.isAnthropic()) {
      const p = await this.getAnthropicProvider();
      yield* p.chatStream(options);
      return;
    }
    if (this.isGemini()) {
      const p = await this.getGeminiProvider();
      yield* p.chatStream(options);
      return;
    }
    if (this.isNvidia()) {
      const p = await this.getNvidiaProvider();
      yield* p.chatStream(options);
      return;
    }
    yield* super.chatStream(options);
  }
}
