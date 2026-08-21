import { OpenAICompatibleProvider } from './OpenAICompatibleProvider';
import { AIRequestOptions } from './types';

export class OpenRouterProvider extends OpenAICompatibleProvider {
  readonly id = 'openrouter';
  readonly name = 'OpenRouter';
  readonly baseUrl = 'https://openrouter.ai/api/v1';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      // Meta Llama
      { id: 'meta-llama/llama-3.3-70b-instruct',     name: 'Llama 3.3 70B' },
      { id: 'meta-llama/llama-4-scout',               name: 'Llama 4 Scout' },
      { id: 'meta-llama/llama-4-maverick',            name: 'Llama 4 Maverick' },
      // Anthropic
      { id: 'anthropic/claude-opus-4',                name: 'Claude Opus 4' },
      { id: 'anthropic/claude-sonnet-4',              name: 'Claude Sonnet 4' },
      { id: 'anthropic/claude-3.5-sonnet',            name: 'Claude 3.5 Sonnet' },
      { id: 'anthropic/claude-3.5-haiku',             name: 'Claude 3.5 Haiku' },
      // Google
      { id: 'google/gemini-2.5-pro',                  name: 'Gemini 2.5 Pro' },
      { id: 'google/gemini-2.5-flash',                name: 'Gemini 2.5 Flash' },
      { id: 'google/gemini-2.0-flash-001',            name: 'Gemini 2.0 Flash' },
      // OpenAI
      { id: 'openai/gpt-4.1',                         name: 'GPT-4.1' },
      { id: 'openai/gpt-4o',                          name: 'GPT-4o' },
      { id: 'openai/o4-mini',                         name: 'o4-mini' },
      { id: 'openai/o3',                              name: 'o3' },
      // DeepSeek
      { id: 'deepseek/deepseek-r1',                   name: 'DeepSeek R1' },
      { id: 'deepseek/deepseek-chat-v3-0324',         name: 'DeepSeek V3' },
      // Mistral
      { id: 'mistralai/mistral-large-2411',           name: 'Mistral Large' },
      { id: 'mistralai/codestral-2501',               name: 'Codestral' },
      // Qwen
      { id: 'qwen/qwen3-235b-a22b',                   name: 'Qwen3 235B' },
      { id: 'qwen/qwen-2.5-72b-instruct',             name: 'Qwen 2.5 72B' },
      // Moonshot
      { id: 'moonshotai/kimi-k2',                     name: 'Kimi K2' },
      // Microsoft
      { id: 'microsoft/phi-4',                        name: 'Phi-4' },
    ];
  }

  /** OpenRouter requires HTTP-Referer + X-Title headers for attribution. */
  protected buildHeaders(_options: AIRequestOptions): Record<string, string> {
    return {
      'Authorization': `Bearer ${this.apiKey}`,
      'Content-Type': 'application/json',
      'HTTP-Referer': 'https://github.com/SuprakashBakali/office-ai-copilot',
      'X-Title': 'Office AI Copilot',
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
