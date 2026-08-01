import { OpenAICompatibleProvider } from './OpenAICompatibleProvider';

export class NvidiaProvider extends OpenAICompatibleProvider {
  readonly id = 'nvidia';
  readonly name = 'NVIDIA NIM';
  readonly baseUrl = 'https://integrate.api.nvidia.com/v1';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      // Llama 3.x
      { id: 'meta/llama-3.1-405b-instruct', name: 'Llama 3.1 405B' },
      { id: 'meta/llama-3.1-70b-instruct', name: 'Llama 3.1 70B' },
      { id: 'meta/llama-3.1-8b-instruct', name: 'Llama 3.1 8B' },
      { id: 'meta/llama-3.2-3b-instruct', name: 'Llama 3.2 3B' },
      { id: 'meta/llama-3.2-1b-instruct', name: 'Llama 3.2 1B' },
      { id: 'meta/llama-3.3-70b-instruct', name: 'Llama 3.3 70B' },
      // NVIDIA Nemotron
      { id: 'nvidia/nemotron-4-340b-instruct', name: 'Nemotron-4 340B' },
      { id: 'nvidia/llama-3.1-nemotron-70b-instruct', name: 'Nemotron 70B' },
      { id: 'nvidia/llama-3.1-nemotron-51b-instruct', name: 'Nemotron 51B' },
      // Mistral / Mixtral
      { id: 'mistralai/mixtral-8x7b-instruct-v0.1', name: 'Mixtral 8x7B' },
      { id: 'mistralai/mixtral-8x22b-instruct-v0.1', name: 'Mixtral 8x22B' },
      { id: 'mistralai/mistral-7b-instruct-v0.3', name: 'Mistral 7B' },
      { id: 'mistralai/mistral-large-2-instruct', name: 'Mistral Large 2' },
      // DeepSeek
      { id: 'deepseek-ai/deepseek-r1', name: 'DeepSeek R1' },
      { id: 'deepseek-ai/deepseek-r1-distill-llama-70b', name: 'DeepSeek R1 Distill 70B' },
      { id: 'deepseek-ai/deepseek-r1-distill-qwen-7b', name: 'DeepSeek R1 Distill 7B' },
      { id: 'deepseek-ai/deepseek-coder-6.7b-instruct', name: 'DeepSeek Coder 6.7B' },
      // Qwen
      { id: 'qwen/qwen2.5-72b-instruct', name: 'Qwen 2.5 72B' },
      { id: 'qwen/qwen2.5-7b-instruct', name: 'Qwen 2.5 7B' },
      { id: 'qwen/qwen2-7b-instruct', name: 'Qwen 2 7B' },
      // Google Gemma
      { id: 'google/gemma-2-27b-it', name: 'Gemma 2 27B' },
      { id: 'google/gemma-2-9b-it', name: 'Gemma 2 9B' },
      { id: 'google/gemma-2-2b-it', name: 'Gemma 2 2B' },
      // Microsoft Phi
      { id: 'microsoft/phi-3-medium-128k-instruct', name: 'Phi-3 Medium' },
      { id: 'microsoft/phi-3-mini-128k-instruct', name: 'Phi-3 Mini' },
      { id: 'microsoft/phi-3.5-mini-instruct', name: 'Phi-3.5 Mini' },
      // Code models
      { id: 'meta/codellama-70b-instruct', name: 'CodeLlama 70B' },
      // Custom model placeholder — resolved by ModelSelector input
      { id: '__custom__', name: '✏️ Custom Model ID...' },
    ];
  }
}
