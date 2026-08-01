import { OpenAICompatibleProvider } from './OpenAICompatibleProvider';

export class OpenAIProvider extends OpenAICompatibleProvider {
  readonly id = 'openai';
  readonly name = 'OpenAI';
  readonly baseUrl = 'https://api.openai.com/v1';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      { id: 'gpt-4o', name: 'GPT-4o' },
      { id: 'gpt-4o-mini', name: 'GPT-4o Mini' },
      { id: 'gpt-4-turbo', name: 'GPT-4 Turbo' },
      { id: 'gpt-3.5-turbo', name: 'GPT-3.5 Turbo' },
    ];
  }
}
