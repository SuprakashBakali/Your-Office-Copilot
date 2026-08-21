import { OpenAICompatibleProvider } from './OpenAICompatibleProvider';
import { AIRequestOptions } from './types';

export class GroqProvider extends OpenAICompatibleProvider {
  readonly id = 'groq';
  readonly name = 'Groq';
  readonly baseUrl = 'https://api.groq.com/openai/v1';

  /** Groq rejects stream_options — disable it. */
  protected supportsStreamOptions = false;

  requiresKey(): boolean { return true; }


  getModels() {
    return [
      // Llama 3.x — fast inference via Groq
      { id: 'llama-3.3-70b-versatile',              name: 'Llama 3.3 70B' },
      { id: 'llama-3.1-70b-versatile',              name: 'Llama 3.1 70B' },
      { id: 'llama-3.1-8b-instant',                 name: 'Llama 3.1 8B (instant)' },
      // Meta Llama 4
      { id: 'meta-llama/llama-4-scout-17b-16e-instruct', name: 'Llama 4 Scout 17B' },
      { id: 'meta-llama/llama-4-maverick-17b-128e-instruct', name: 'Llama 4 Maverick 17B' },
      // DeepSeek R1 distills
      { id: 'deepseek-r1-distill-llama-70b',        name: 'DeepSeek R1 Distill 70B' },
      { id: 'deepseek-r1-distill-qwen-32b',         name: 'DeepSeek R1 Distill Qwen 32B' },
      // Qwen
      { id: 'qwen-qwen2.5-72b',                     name: 'Qwen 2.5 72B' },
      { id: 'qwen-qwen2.5-coder-32b',               name: 'Qwen 2.5 Coder 32B' },
      // Mixtral
      { id: 'mixtral-8x7b-32768',                   name: 'Mixtral 8x7B' },
      // Gemma
      { id: 'gemma2-9b-it',                         name: 'Gemma 2 9B' },
      // Moonshot Kimi
      { id: 'moonshotai/kimi-k2-instruct',          name: 'Kimi K2' },
    ];
  }

  /** Groq does not accept temperature > 2.0 and rejects stream_options.
   *  Also strip reasoning_effort / any extra fields that cause 400s. */
  protected buildBody(options: AIRequestOptions): Record<string, unknown> {
    const body = super.buildBody(options);
    // Clamp temperature to Groq's accepted range [0, 2]
    if (typeof body['temperature'] === 'number') {
      body['temperature'] = Math.min(body['temperature'] as number, 2.0);
    }
    // Groq rejects stream_options — it's added in buildRequest, handled via deletion
    return body;
  }
}
