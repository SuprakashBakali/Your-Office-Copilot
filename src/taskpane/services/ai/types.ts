export interface AIRequestOptions {
  model: string;
  messages: Array<{role: 'system'|'user'|'assistant'; content: string}>;
  temperature?: number;
  maxTokens?: number;
  stream?: boolean;
}

export interface AIStreamChunk {
  content: string;
  done: boolean;
}

export interface AIResponse {
  content: string;
  model: string;
  tokens?: { prompt: number; completion: number; total: number };
}

export abstract class BaseAIProvider {
  abstract readonly id: string;
  abstract readonly name: string;
  abstract readonly baseUrl: string;
  
  protected apiKey: string = '';
  
  setApiKey(key: string): void { this.apiKey = key; }
  getApiKey(): string { return this.apiKey; }
  isConfigured(): boolean { return !!this.apiKey || !this.requiresKey(); }
  abstract requiresKey(): boolean;
  abstract getModels(): Array<{id: string; name: string}>;
  abstract chat(options: AIRequestOptions): Promise<AIResponse>;
  abstract chatStream(options: AIRequestOptions): AsyncGenerator<AIStreamChunk>;
}

// Utility for parsing OpenAI-compatible SSE streams
export async function* parseOpenAISSEStream(response: Response): AsyncGenerator<AIStreamChunk> {
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

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || !trimmed.startsWith("data: ")) continue;
        
        const data = trimmed.slice(6);
        if (data === "[DONE]") {
          yield { content: "", done: true };
          return;
        }

        try {
          const parsed = JSON.parse(data);
          const content = parsed.choices?.[0]?.delta?.content;
          if (content) {
            yield { content, done: false };
          }
        } catch (e) {
          console.error("Failed to parse SSE chunk", e);
        }
      }
    }
  } finally {
    reader.releaseLock();
  }
}
