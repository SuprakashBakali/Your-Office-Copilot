import { BaseAIProvider, AIRequestOptions, AIResponse, AIStreamChunk, parseOpenAISSEStream } from './types';

export class NvidiaProvider extends BaseAIProvider {
  readonly id = 'nvidia';
  readonly name = 'NVIDIA NIM';
  readonly baseUrl = 'https://integrate.api.nvidia.com/v1';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      { id: 'meta/llama-3.1-70b-instruct', name: 'Llama 3.1 70B' },
      { id: 'meta/llama-3.1-8b-instruct', name: 'Llama 3.1 8B' },
      { id: 'mistralai/mixtral-8x7b-instruct-v0.1', name: 'Mixtral 8x7B' },
      { id: 'nvidia/nemotron-4-340b-instruct', name: 'Nemotron-4 340B' },
      { id: 'microsoft/phi-3-medium-128k-instruct', name: 'Phi-3 Medium' }
    ];
  }

  private getHeaders() {
    return {
      'Authorization': `Bearer ${this.apiKey}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    };
  }

  async chat(options: AIRequestOptions): Promise<AIResponse> {
    const response = await fetch(`${this.baseUrl}/chat/completions`, {
      method: 'POST',
      headers: this.getHeaders(),
      body: JSON.stringify({
        model: options.model,
        messages: options.messages,
        temperature: options.temperature,
        max_tokens: options.maxTokens ?? 1024,
        stream: false
      })
    });

    if (!response.ok) {
      const errBody = await response.text().catch(() => '');
      const errMsg = errBody ? `${response.status} — ${errBody.substring(0, 200)}` : `${response.status} ${response.statusText}`;
      throw new Error(`NVIDIA API Error: ${errMsg}`);
    }

    const data = await response.json();
    return {
      content: data.choices[0].message.content,
      model: data.model,
      tokens: {
        prompt: data.usage?.prompt_tokens ?? 0,
        completion: data.usage?.completion_tokens ?? 0,
        total: data.usage?.total_tokens ?? 0
      }
    };
  }

  async *chatStream(options: AIRequestOptions): AsyncGenerator<AIStreamChunk> {
    const response = await fetch(`${this.baseUrl}/chat/completions`, {
      method: 'POST',
      headers: {
        ...this.getHeaders(),
        'Accept': 'text/event-stream'
      },
      body: JSON.stringify({
        model: options.model,
        messages: options.messages,
        temperature: options.temperature,
        max_tokens: options.maxTokens ?? 1024,
        stream: true
      })
    });

    if (!response.ok) {
      const errBody = await response.text().catch(() => '');
      const errMsg = errBody ? `${response.status} — ${errBody.substring(0, 200)}` : `${response.status} ${response.statusText}`;
      throw new Error(`NVIDIA API Error: ${errMsg}`);
    }

    yield* parseOpenAISSEStream(response);
  }
}
