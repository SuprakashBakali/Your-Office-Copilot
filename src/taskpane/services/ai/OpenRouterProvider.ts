import { OpenAICompatibleProvider } from './OpenAICompatibleProvider';
import { AIRequestOptions } from './types';

export class OpenRouterProvider extends OpenAICompatibleProvider {
  readonly id = 'openrouter';
  readonly name = 'OpenRouter';
  readonly baseUrl = 'https://openrouter.ai/api/v1';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      { id: 'meta-llama/llama-3.1-70b-instruct', name: 'Llama 3.1 70B' },
      { id: 'anthropic/claude-sonnet-4-20250514', name: 'Claude 3.5 Sonnet' },
      { id: 'google/gemini-2.5-flash', name: 'Gemini 2.5 Flash' },
      { id: 'openai/gpt-4o', name: 'GPT-4o' },
    ];
  }

  /** OpenRouter requires HTTP-Referer + X-Title headers for attribution. */
  protected buildHeaders(_options: AIRequestOptions): Record<string, string> {
    return {
      'Authorization': `Bearer ${this.apiKey}`,
      'Content-Type': 'application/json',
      'HTTP-Referer': typeof window !== 'undefined' ? window.location.href : 'http://localhost:3000',
      'X-Title': 'Your Co-Pilot',
    };
  }

  /** OpenRouter supports a `plugins` array for web search. */
  protected buildBody(options: AIRequestOptions): Record<string, unknown> {
    const body = super.buildBody(options);
    if (options.webSearch) {
      body['plugins'] = [{ id: 'web', max_results: 5 }];
    }
    return body;
  }
}
