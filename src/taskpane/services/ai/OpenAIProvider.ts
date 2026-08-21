import { OpenAICompatibleProvider } from './OpenAICompatibleProvider';

export class OpenAIProvider extends OpenAICompatibleProvider {
  readonly id = 'openai';
  readonly name = 'OpenAI';
  readonly baseUrl = 'https://api.openai.com/v1';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      // GPT-4.1 series (2025)
      { id: 'gpt-4.1',          name: 'GPT-4.1' },
      { id: 'gpt-4.1-mini',     name: 'GPT-4.1 Mini' },
      { id: 'gpt-4.1-nano',     name: 'GPT-4.1 Nano' },
      // GPT-4o series
      { id: 'gpt-4o',           name: 'GPT-4o' },
      { id: 'gpt-4o-mini',      name: 'GPT-4o Mini' },
      // o-series reasoning models
      { id: 'o4-mini',          name: 'o4-mini (reasoning)' },
      { id: 'o3',               name: 'o3 (reasoning)' },
      { id: 'o3-mini',          name: 'o3-mini (reasoning)' },
      { id: 'o1',               name: 'o1 (reasoning)' },
      { id: 'o1-mini',          name: 'o1-mini (reasoning)' },
      // Legacy
      { id: 'gpt-4-turbo',      name: 'GPT-4 Turbo' },
      { id: 'gpt-3.5-turbo',    name: 'GPT-3.5 Turbo' },
    ];
  }
}
