/**
 * GenericOpenAIProvider — works with any user-supplied OpenAI-compatible API
 * endpoint. Used by the "My Models" settings panel for custom provider
 * configurations (custom baseUrl + key + modelId).
 */
import { OpenAICompatibleProvider } from './OpenAICompatibleProvider';
import { AIRequestOptions } from './types';

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

  /** OpenRouter attribution headers when the user points a custom model at
   *  openrouter.ai. */
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

  /** OpenRouter web-search plugin support (only applies to openrouter.ai). */
  protected buildBody(options: AIRequestOptions): Record<string, unknown> {
    const body = super.buildBody(options);
    if (options.webSearch && this._baseUrl.includes('openrouter.ai')) {
      body['plugins'] = [{ id: 'web', max_results: 5 }];
    }
    return body;
  }
}
