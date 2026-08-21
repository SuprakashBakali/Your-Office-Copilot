import { OpenAICompatibleProvider } from './OpenAICompatibleProvider';

export class NvidiaProvider extends OpenAICompatibleProvider {
  readonly id = 'nvidia';
  readonly name = 'NVIDIA NIM';
  readonly baseUrl = 'https://integrate.api.nvidia.com/v1';

  /** NVIDIA NIM does not support stream_options. */
  protected supportsStreamOptions = false;

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      // Meta Llama 4 (2025)
      { id: 'meta/llama-4-scout-17b-16e-instruct',    name: 'Llama 4 Scout 17B' },
      { id: 'meta/llama-4-maverick-17b-128e-instruct', name: 'Llama 4 Maverick 17B' },
      // Llama 3.x
      { id: 'meta/llama-3.3-70b-instruct',            name: 'Llama 3.3 70B' },
      { id: 'meta/llama-3.1-405b-instruct',           name: 'Llama 3.1 405B' },
      { id: 'meta/llama-3.1-70b-instruct',            name: 'Llama 3.1 70B' },
      { id: 'meta/llama-3.1-8b-instruct',             name: 'Llama 3.1 8B' },
      // NVIDIA Nemotron
      { id: 'nvidia/llama-3.1-nemotron-ultra-253b-v1', name: 'Nemotron Ultra 253B' },
      { id: 'nvidia/llama-3.1-nemotron-70b-instruct', name: 'Nemotron 70B' },
      // DeepSeek
      { id: 'deepseek-ai/deepseek-r1',                name: 'DeepSeek R1' },
      { id: 'deepseek-ai/deepseek-r1-distill-llama-70b', name: 'DeepSeek R1 Distill 70B' },
      // Mistral
      { id: 'mistralai/mistral-large-2-instruct',     name: 'Mistral Large 2' },
      { id: 'mistralai/mixtral-8x22b-instruct-v0.1',  name: 'Mixtral 8x22B' },
      // Qwen
      { id: 'qwen/qwen2.5-72b-instruct',              name: 'Qwen 2.5 72B' },
      { id: 'qwen/qwen2.5-coder-32b-instruct',        name: 'Qwen 2.5 Coder 32B' },
      // Microsoft Phi
      { id: 'microsoft/phi-4',                        name: 'Phi-4' },
      { id: 'microsoft/phi-3.5-mini-instruct',        name: 'Phi-3.5 Mini' },
      // Google
      { id: 'google/gemma-3-27b-it',                  name: 'Gemma 3 27B' },
      // Custom model placeholder
      { id: '__custom__', name: '✏️ Custom Model ID...' },
    ];
  }
}
