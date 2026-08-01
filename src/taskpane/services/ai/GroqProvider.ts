import { OpenAICompatibleProvider } from './OpenAICompatibleProvider';

export class GroqProvider extends OpenAICompatibleProvider {
  readonly id = 'groq';
  readonly name = 'Groq';
  readonly baseUrl = 'https://api.groq.com/openai/v1';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      { id: 'llama-3.1-70b-versatile', name: 'Llama 3.1 70B' },
      { id: 'llama-3.1-8b-instant', name: 'Llama 3.1 8B' },
      { id: 'mixtral-8x7b-32768', name: 'Mixtral 8x7B' },
      { id: 'gemma2-9b-it', name: 'Gemma 2 9B' },
    ];
  }
}
