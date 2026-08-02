import { useState, useCallback, useRef } from 'react';
import { AIProviderType, ChatMessage } from '../types';
import { useSettings } from './useSettings';
import { getProvider } from '../services/ai/ProviderFactory';
import { GenericOpenAIProvider } from '../services/ai/GenericOpenAIProvider';
import { AIRequestOptions } from '../services/ai/types';

export function useAI() {
  const [isStreaming, setIsStreaming] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const { settings } = useSettings();
  const abortControllerRef = useRef<AbortController | null>(null);

  /** Resolve the provider + API key + model id from the active custom model
   *  (if any) or fall back to the active built-in provider. */
  const resolveProvider = useCallback(() => {
    const activeCustomModel = (settings.customModels || []).find(
      m => m.id === settings.activeCustomModelId,
    );
    const apiKey = activeCustomModel?.apiKey || settings.apiKeys[settings.activeProvider] || '';
    const modelId = activeCustomModel?.modelId ?? settings.activeModel;

    const provider = activeCustomModel?.baseUrl
      ? new GenericOpenAIProvider(activeCustomModel.baseUrl, apiKey)
      : getProvider(activeCustomModel?.provider ?? settings.activeProvider);

    if (provider.requiresKey() && !apiKey) {
      throw new Error(`No API key — add one in Settings → My Models.`);
    }
    if (apiKey && !activeCustomModel?.baseUrl) provider.setApiKey(apiKey);

    return { provider, modelId };
  }, [settings]);

  const buildRequestOptions = useCallback(
    (messages: ChatMessage[], options: any, stream: boolean, signal?: AbortSignal): AIRequestOptions => {
      const { modelId } = resolveProvider();
      return {
        model: modelId,
        messages: messages.map(m => ({
          role: m.role as 'system' | 'user' | 'assistant',
          content: m.content,
        })),
        temperature: 0.7,
        maxTokens: 2048,
        stream,
        signal,
        ...options,
      };
    },
    [resolveProvider],
  );

  const sendMessage = useCallback(
    async (messages: ChatMessage[], options?: any): Promise<string> => {
      try {
        setIsStreaming(true);
        setError(null);
        const { provider } = resolveProvider();
        const requestOptions = buildRequestOptions(messages, options, false);
        const response = await provider.chat(requestOptions);
        setIsStreaming(false);
        return response.content;
      } catch (err) {
        setIsStreaming(false);
        const msg = (err as Error).message;
        // Don't surface "aborted" as an error — it's a user action.
        if (msg !== 'Stream cancelled' && !(err as Error).name?.includes('Abort')) {
          setError(msg);
        }
        throw err;
      }
    },
    [resolveProvider, buildRequestOptions],
  );

  const sendMessageStream = useCallback(
    async (
      messages: ChatMessage[],
      options: any,
      onChunk: (chunk: string) => void,
      onThinking?: (chunk: string) => void,
    ): Promise<{ text: string; thinking: string }> => {
      try {
        setIsStreaming(true);
        setError(null);
        abortControllerRef.current = new AbortController();

        const { provider } = resolveProvider();
        const requestOptions = buildRequestOptions(
          messages,
          options,
          true,
          abortControllerRef.current.signal,
        );

        let fullText = '';
        let fullThinking = '';
        const stream = provider.chatStream(requestOptions);

        for await (const chunk of stream) {
          if (abortControllerRef.current?.signal.aborted) {
            break;
          }
          if (chunk.thinking) {
            fullThinking += chunk.thinking;
            onThinking?.(chunk.thinking);
          }
          if (chunk.content) {
            fullText += chunk.content;
            onChunk(chunk.content);
          }
          if (chunk.done) break;
        }

        setIsStreaming(false);
        abortControllerRef.current = null;
        return { text: fullText, thinking: fullThinking };
      } catch (err) {
        setIsStreaming(false);
        abortControllerRef.current = null;
        const msg = (err as Error).message;
        if (msg !== 'Stream cancelled' && !(err as Error).name?.includes('Abort')) {
          setError(msg);
        }
        return { text: '', thinking: '' };
      }
    },
    [resolveProvider, buildRequestOptions],
  );

  const cancelStream = useCallback(() => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
    }
  }, []);

  const switchProvider = useCallback((_provider: AIProviderType) => {
    // Provider switching is handled by settings update
  }, []);

  const switchModel = useCallback((_model: string) => {
    // Model switching is handled by settings update
  }, []);

  const testConnection = useCallback(
    async (provider: AIProviderType, apiKey: string): Promise<{ ok: boolean; error?: string }> => {
      try {
        const p = getProvider(provider);
        // Save the original key so we can restore it after the test.
        // ProviderFactory stores providers as SINGLETONS — if we mutate the
        // key here and don't restore it, subsequent real requests would use
        // the test key instead of the user's actual key.
        const originalKey = p.getApiKey();
        p.setApiKey(apiKey);
        try {
          const models = p.getModels().filter(m => m.id !== '__custom__');
          const testModel =
            models.find(m => m.id.includes('8b') || m.id.includes('7b') || m.id.includes('mini') || m.id.includes('flash') || m.id.includes('haiku')) ||
            models[models.length - 2] ||
            models[0];
          const response = await p.chat({
            model: testModel?.id || '',
            messages: [{ role: 'user', content: 'Hi' }],
            maxTokens: 5,
            stream: false,
          });
          return { ok: !!response.content };
        } finally {
          // Restore the original key so the singleton isn't left mutated.
          p.setApiKey(originalKey);
        }
      } catch (err) {
        return { ok: false, error: (err as Error).message };
      }
    },
    [],
  );

  return {
    sendMessage,
    sendMessageStream,
    cancelStream,
    isStreaming,
    error,
    switchProvider,
    switchModel,
    testConnection,
  };
}
