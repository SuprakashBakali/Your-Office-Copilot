import { BaseAIProvider, AIRequestOptions, AIResponse, AIStreamChunk } from './types';

const PROXY_URL = '/api/proxy';

export class AnthropicProvider extends BaseAIProvider {
  readonly id = 'anthropic';
  readonly name = 'Anthropic Claude';
  readonly baseUrl = 'https://api.anthropic.com/v1';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      { id: 'claude-sonnet-4-20250514', name: 'Claude 3.5 Sonnet' },
      { id: 'claude-3-5-haiku-20241022', name: 'Claude 3.5 Haiku' },
      { id: 'claude-3-opus-20240229', name: 'Claude 3 Opus' }
    ];
  }

  private getHeaders() {
    return {
      'x-api-key': this.apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
      'Content-Type': 'application/json'
    };
  }

  private formatMessages(messages: AIRequestOptions['messages']) {
    let systemPrompt = '';
    const formattedMessages = [];
    
    for (const msg of messages) {
      if (msg.role === 'system') {
        systemPrompt += msg.content + '\n';
      } else {
        formattedMessages.push({
          role: msg.role === 'user' ? 'user' : 'assistant',
          content: msg.content
        });
      }
    }
    
    return { system: systemPrompt.trim() || undefined, messages: formattedMessages };
  }

  async chat(options: AIRequestOptions): Promise<AIResponse> {
    const { system, messages } = this.formatMessages(options.messages);

    const response = await fetch(PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal: options.signal,
      body: JSON.stringify({
        targetUrl: `${this.baseUrl}/messages`,
        headers: this.getHeaders(),
        body: {
          model: options.model,
          messages: messages,
          system: system,
          temperature: options.temperature,
          max_tokens: options.maxTokens ?? 1024,
          stream: false
        }
      })
    });

    if (!response.ok) {
      const errBody = await response.text().catch(() => '');
      const errMsg = errBody ? `${response.status} — ${errBody.substring(0, 200)}` : `${response.status} ${response.statusText}`;
      throw new Error(`Anthropic API Error: ${errMsg}`);
    }

    const data = await response.json();
    return {
      content: data.content?.[0]?.text ?? '',
      model: data.model,
      tokens: {
        prompt: data.usage?.input_tokens ?? 0,
        completion: data.usage?.output_tokens ?? 0,
        total: (data.usage?.input_tokens ?? 0) + (data.usage?.output_tokens ?? 0)
      }
    };
  }

  async *chatStream(options: AIRequestOptions): AsyncGenerator<AIStreamChunk> {
    const { system, messages } = this.formatMessages(options.messages);

    const response = await fetch(PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal: options.signal,
      body: JSON.stringify({
        targetUrl: `${this.baseUrl}/messages`,
        headers: {
          ...this.getHeaders(),
          'Accept': 'text/event-stream'
        },
        body: {
          model: options.model,
          messages: messages,
          system: system,
          temperature: options.temperature,
          max_tokens: options.maxTokens ?? 1024,
          stream: true
        }
      })
    });

    if (!response.ok) {
      const errBody = await response.text().catch(() => '');
      const errMsg = errBody ? `${response.status} — ${errBody.substring(0, 200)}` : `${response.status} ${response.statusText}`;
      throw new Error(`Anthropic API Error: ${errMsg}`);
    }

    if (!response.body) throw new Error("No response body");
    
    const reader = response.body.getReader();
    const decoder = new TextDecoder("utf-8");
    let buffer = "";

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || "";

        let eventType = '';

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed) continue;
          
          if (trimmed.startsWith('event: ')) {
            eventType = trimmed.slice(7);
          } else if (trimmed.startsWith('data: ')) {
            const dataStr = trimmed.slice(6);
            if (dataStr === '[DONE]') {
              yield { content: '', done: true };
              return;
            }
            
            try {
              const data = JSON.parse(dataStr);
              if (eventType === 'content_block_delta' && data.delta?.type === 'text_delta') {
                yield { content: data.delta.text, done: false };
              } else if (eventType === 'message_stop') {
                yield { content: '', done: true };
                return;
              }
            } catch (e) {
              console.error("Failed to parse Anthropic SSE chunk", e);
            }
          }
        }
      }
    } finally {
      reader.releaseLock();
    }
  }
}
