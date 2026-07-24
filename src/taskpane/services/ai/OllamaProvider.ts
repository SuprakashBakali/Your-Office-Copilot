import { BaseAIProvider, AIRequestOptions, AIResponse, AIStreamChunk, parseOpenAISSEStream } from './types';

export class OllamaProvider extends BaseAIProvider {
  readonly id = 'ollama';
  readonly name = 'Ollama (Local)';
  readonly baseUrl = 'http://localhost:11434/v1';

  requiresKey(): boolean { return false; }
  
  private dynamicModels: Array<{id: string, name: string}> = [];

  constructor() {
    super();
    this.fetchLocalModels().catch(console.error);
  }

  private async fetchLocalModels() {
    try {
      const res = await fetch('http://localhost:11434/api/tags');
      if (res.ok) {
        const data = await res.json();
        if (data.models && Array.isArray(data.models)) {
          this.dynamicModels = data.models.map((m: any) => ({
            id: m.name,
            name: m.name
          }));
        }
      }
    } catch (e) {
      console.warn("Could not fetch Ollama models, using fallback list.", e);
    }
  }

  getModels() {
    if (this.dynamicModels.length > 0) {
      return this.dynamicModels;
    }
    return [
      { id: 'llama3.1', name: 'Llama 3.1' },
      { id: 'mistral', name: 'Mistral' },
      { id: 'phi3', name: 'Phi-3' },
      { id: 'codellama', name: 'Code Llama' }
    ];
  }

  private getHeaders() {
    return {
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
        max_tokens: options.maxTokens,
        stream: false
      })
    });

    if (!response.ok) {
      const errBody = await response.text().catch(() => '');
      const errMsg = errBody ? `${response.status} — ${errBody.substring(0, 200)}` : `${response.status} ${response.statusText}`;
      throw new Error(`Ollama API Error: ${errMsg}`);
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
        max_tokens: options.maxTokens,
        stream: true
      })
    });

    if (!response.ok) {
      const errBody = await response.text().catch(() => '');
      const errMsg = errBody ? `${response.status} — ${errBody.substring(0, 200)}` : `${response.status} ${response.statusText}`;
      throw new Error(`Ollama API Error: ${errMsg}`);
    }

    yield* parseOpenAISSEStream(response);
  }
}
