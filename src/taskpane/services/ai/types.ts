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
//
// Thinking-block handling: DeepSeek-R1 and similar reasoning models stream
// their chain-of-thought either as:
//   (a) a separate `delta.reasoning_content` field, OR
//   (b) inline in `delta.content` wrapped in <think>...</think> tags.
// Because streaming fragments the tags across chunks (e.g. one chunk is
// "<think>", the next is "Let me analyze", the next is "</think>"), we
// track an `inThinkingBlock` state to route ALL content between the tags
// to the `thinking` channel — not just chunks that happen to start with
// `<think>`.
export async function* parseOpenAISSEStream(
  response: Response,
  signal?: AbortSignal,
): AsyncGenerator<AIStreamChunk> {
  if (!response.body) throw new Error("No response body");

  const reader = response.body.getReader();
  const decoder = new TextDecoder("utf-8");
  let buffer = "";
  let inThinkingBlock = false;

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
          const content: string | undefined = delta?.content;

          // Case (a): explicit reasoning_content field (DeepSeek-R1's API).
          if (delta?.reasoning_content) {
            yield { content: '', done: false, thinking: delta.reasoning_content };
            continue;
          }

          // Case (b): inline <think>...</think> blocks in the content stream.
          if (typeof content === 'string' && content.length > 0) {
            // Split on <think> and </think> tags, preserving them.
            // This handles cases where the tag and content arrive in the
            // same chunk OR split across chunks.
            const parts = content.split(/(<\/?think>)/);
            for (const part of parts) {
              if (part === '<think>') {
                inThinkingBlock = true;
                yield { content: '', done: false, thinking: '<think>' };
              } else if (part === '</think>') {
                inThinkingBlock = false;
                yield { content: '', done: false, thinking: '</think>' };
              } else if (part.length > 0) {
                if (inThinkingBlock) {
                  yield { content: '', done: false, thinking: part };
                } else {
                  yield { content: part, done: false };
                }
              }
            }
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
