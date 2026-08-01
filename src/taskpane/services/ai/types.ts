export interface AIRequestOptions {
  model: string;
  messages: Array<{role: 'system'|'user'|'assistant'; content: string}>;
  temperature?: number;
  maxTokens?: number;
  stream?: boolean;
  webSearch?: boolean;
  /** Optional AbortSignal — when aborted, the underlying fetch is cancelled. */
  signal?: AbortSignal;
}

export interface AIStreamChunk {
  content: string;
  done: boolean;
  /** Optional reasoning/thinking token from models that stream it separately
   *  (e.g. DeepSeek-R1 `<think>...</think>` blocks). */
  thinking?: string;
}

export interface AIResponse {
  content: string;
  model: string;
  tokens?: { prompt: number; completion: number; total: number };
  /** Optional reasoning/thinking trace from models that expose it
   *  (e.g. DeepSeek-R1, Claude w/ extended thinking, OpenAI o1). */
  thinking?: string;
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

// Utility for parsing OpenAI-compatible SSE streams.
// If an AbortSignal is provided and becomes aborted, the reader is cancelled
// and the generator returns cleanly.
export async function* parseOpenAISSEStream(
  response: Response,
  signal?: AbortSignal,
): AsyncGenerator<AIStreamChunk> {
  if (!response.body) throw new Error("No response body");

  const reader = response.body.getReader();
  const decoder = new TextDecoder("utf-8");
  let buffer = "";

  const onAbort = () => { reader.cancel().catch(() => {}); };
  if (signal) {
    if (signal.aborted) { await reader.cancel().catch(() => {}); return; }
    signal.addEventListener('abort', onAbort, { once: true });
  }

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
          const delta = parsed.choices?.[0]?.delta;
          const content = delta?.content;
          // DeepSeek-R1 and similar reasoning models expose thinking either as
          // a separate `reasoning_content` field OR inline wrapped in <think>.
          // We capture both so the UI can render a collapsible thinking block.
          const thinking =
            delta?.reasoning_content ??
            (typeof content === 'string' && content.startsWith('<think>')
              ? content
              : undefined);
          if (content) {
            yield { content, done: false, thinking };
          } else if (thinking) {
            yield { content: '', done: false, thinking };
          }
        } catch (e) {
          console.error("Failed to parse SSE chunk", e);
        }
      }
    }
  } finally {
    if (signal) signal.removeEventListener('abort', onAbort);
    reader.releaseLock();
  }
}
