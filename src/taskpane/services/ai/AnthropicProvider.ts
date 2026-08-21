import { BaseAIProvider, AIRequestOptions, AIResponse, AIStreamChunk } from './types';

const PROXY_URL = '/api/proxy';

/** Gemini 2.5 Flash/Pro support thinking — enable it by default for those. */
const THINKING_MODELS = ['gemini-2.5', 'gemini-exp'];
function supportsThinking(modelId: string): boolean {
  return THINKING_MODELS.some(p => modelId.startsWith(p));
}

export class AnthropicProvider extends BaseAIProvider {
  readonly id = 'anthropic';
  readonly name = 'Anthropic Claude';
  readonly baseUrl = 'https://api.anthropic.com/v1';

  requiresKey(): boolean { return true; }

  getModels() {
    return [
      // 2025 flagship
      { id: 'claude-opus-4-20250514',     name: 'Claude Opus 4' },
      { id: 'claude-sonnet-4-20250514',   name: 'Claude Sonnet 4' },
      // Extended thinking
      { id: 'claude-3-7-sonnet-20250219', name: 'Claude 3.7 Sonnet (thinking)' },
      // 3.5 series
      { id: 'claude-3-5-sonnet-20241022', name: 'Claude 3.5 Sonnet' },
      { id: 'claude-3-5-haiku-20241022',  name: 'Claude 3.5 Haiku' },
      // Legacy
      { id: 'claude-3-opus-20240229',     name: 'Claude 3 Opus (legacy)' },
    ];
  }

  /** Whether the model supports extended thinking (claude-3-7+). */
  private supportsExtendedThinking(modelId: string): boolean {
    return modelId.includes('3-7') || modelId.includes('opus-4') || modelId.includes('sonnet-4');
  }

  private getHeaders(includeThinkingBeta = false): Record<string, string> {
    const headers: Record<string, string> = {
      'x-api-key': this.apiKey,
      'anthropic-version': '2023-06-01',
      'Content-Type': 'application/json',
    };
    if (includeThinkingBeta) {
      headers['anthropic-beta'] = 'interleaved-thinking-2025-02-19';
    }
    // anthropic-dangerous-direct-browser-access is only relevant when hitting
    // the Anthropic endpoint directly (not via our proxy). Sending it via
    // proxy is harmless but confusing — omit for clarity.
    return headers;
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
          content: msg.content,
        });
      }
    }

    return { system: systemPrompt.trim() || undefined, messages: formattedMessages };
  }

  /** Build the Anthropic messages API body. */
  private buildBody(options: AIRequestOptions, stream: boolean): Record<string, unknown> {
    const { system, messages } = this.formatMessages(options.messages);
    const useThinking = this.supportsExtendedThinking(options.model);
    const body: Record<string, unknown> = {
      model: options.model,
      messages,
      max_tokens: options.maxTokens ?? 8192,
      stream,
    };
    if (system) body['system'] = system;
    // Don't include temperature when using extended thinking (conflicts with budget_tokens)
    if (!useThinking && options.temperature !== undefined) {
      body['temperature'] = options.temperature;
    }
    if (useThinking) {
      body['thinking'] = { type: 'enabled', budget_tokens: 4096 };
    }
    return body;
  }

  private async throwIfError(response: Response, label = 'Anthropic'): Promise<void> {
    if (response.ok) return;
    let errMsg = `${response.status} ${response.statusText}`;
    try {
      const raw = await response.text();
      if (raw) {
        try {
          const parsed = JSON.parse(raw);
          const msg = parsed?.error?.message || parsed?.message;
          errMsg = msg ? `${response.status} — ${msg}` : `${response.status} — ${raw.substring(0, 300)}`;
        } catch {
          errMsg = `${response.status} — ${raw.substring(0, 300)}`;
        }
      }
    } catch { /* ignore */ }
    throw new Error(`${label} API Error: ${errMsg}`);
  }

  async chat(options: AIRequestOptions): Promise<AIResponse> {
    const useThinking = this.supportsExtendedThinking(options.model);
    const headers = this.getHeaders(useThinking);
    const body = this.buildBody(options, false);

    const response = await fetch(PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal: options.signal,
      body: JSON.stringify({ targetUrl: `${this.baseUrl}/messages`, headers, body }),
    });
    await this.throwIfError(response);

    const data = await response.json();
    // Extract text and thinking blocks from content array
    let content = '';
    let thinking = '';
    for (const block of (data.content ?? [])) {
      if (block.type === 'text') content += block.text ?? '';
      if (block.type === 'thinking') thinking += block.thinking ?? '';
    }
    return {
      content,
      thinking: thinking || undefined,
      model: data.model ?? options.model,
      tokens: {
        prompt: data.usage?.input_tokens ?? 0,
        completion: data.usage?.output_tokens ?? 0,
        total: (data.usage?.input_tokens ?? 0) + (data.usage?.output_tokens ?? 0),
      },
    };
  }

  async *chatStream(options: AIRequestOptions): AsyncGenerator<AIStreamChunk> {
    const useThinking = this.supportsExtendedThinking(options.model);
    const headers = { ...this.getHeaders(useThinking), 'Accept': 'text/event-stream' };
    const body = this.buildBody(options, true);

    const response = await fetch(PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal: options.signal,
      body: JSON.stringify({ targetUrl: `${this.baseUrl}/messages`, headers, body }),
    });
    await this.throwIfError(response);

    if (!response.body) throw new Error('No response body');

    const reader = response.body.getReader();
    const onAbort = () => { reader.cancel().catch(() => {}); };
    if (options.signal) {
      if (options.signal.aborted) { await reader.cancel().catch(() => {}); return; }
      options.signal.addEventListener('abort', onAbort, { once: true });
    }
    const decoder = new TextDecoder('utf-8');
    let buffer = '';
    let eventType = '';
    // Track whether we're inside a thinking content block (block_index → type)
    const blockTypes = new Map<number, string>();

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed) { eventType = ''; continue; }

          if (trimmed.startsWith('event: ')) {
            eventType = trimmed.slice(7).trim();
          } else if (trimmed.startsWith('data: ')) {
            const dataStr = trimmed.slice(6);
            if (dataStr === '[DONE]') {
              yield { content: '', done: true };
              return;
            }
            try {
              const data = JSON.parse(dataStr);

              if (eventType === 'content_block_start') {
                // Register block type so we know if subsequent deltas are thinking
                const idx: number = data.index ?? -1;
                const type: string = data.content_block?.type ?? 'text';
                if (idx >= 0) blockTypes.set(idx, type);
              } else if (eventType === 'content_block_delta') {
                const idx: number = data.index ?? -1;
                const blockType = blockTypes.get(idx) ?? 'text';
                const deltaType: string = data.delta?.type ?? '';

                if (blockType === 'thinking' || deltaType === 'thinking_delta') {
                  // Streaming thinking / chain-of-thought
                  yield { content: '', done: false, thinking: data.delta?.thinking ?? '' };
                } else if (deltaType === 'text_delta') {
                  yield { content: data.delta?.text ?? '', done: false };
                }
              } else if (eventType === 'message_stop') {
                yield { content: '', done: true };
                return;
              }
            } catch {
              // Ignore malformed SSE lines
            }
          }
        }
      }
    } finally {
      if (options.signal) options.signal.removeEventListener('abort', onAbort);
      reader.releaseLock();
    }
  }
}
