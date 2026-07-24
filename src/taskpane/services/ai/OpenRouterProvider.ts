import { BaseAIProvider, AIRequestOptions, AIResponse, AIStreamChunk, parseOpenAISSEStream } from './types';

export class OpenRouterProvider extends BaseAIProvider {
  readonly id = 'openrouter';
  readonly name = 'OpenRouter';
  readonly baseUrl = 'https://openrouter.ai/api/v1';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      { id: 'meta-llama/llama-3.1-70b-instruct', name: 'Llama 3.1 70B' },
      { id: 'anthropic/claude-sonnet-4-20250514', name: 'Claude 3.5 Sonnet' },
      { id: 'google/gemini-2.5-flash', name: 'Gemini 2.5 Flash' },
      { id: 'openai/gpt-4o', name: 'GPT-4o' }
    ];
  }

  private getHeaders() {
    return {
      'Authorization': `Bearer ${this.apiKey}`,
      'HTTP-Referer': window.location.href || 'http://localhost:3000',
      'X-Title': 'Office AI Copilot',
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
      throw new Error(`OpenRouter API Error: ${errMsg}`);
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
      throw new Error(`OpenRouter API Error: ${errMsg}`);
    }

    yield* parseOpenAISSEStream(response);
  }
}
