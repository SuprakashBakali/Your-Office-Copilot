import { useState, useCallback, useRef } from 'react';
import { AIProviderType, ChatMessage } from '../types';
import { loadSettings } from '../utils/storage';
import { getProvider } from '../services/ai/ProviderFactory';
import { GenericOpenAIProvider } from '../services/ai/GenericOpenAIProvider';
import { AIRequestOptions } from '../services/ai/types';

export function useAI() {
  const [isStreaming, setIsStreaming] = useState(false);
  const [currentStreamText, setCurrentStreamText] = useState("");
  const [error, setError] = useState<string | null>(null);
  const abortControllerRef = useRef<AbortController | null>(null);
  const streamTextRef = useRef<string>("");

  const sendMessage = useCallback(async (messages: ChatMessage[], options?: any): Promise<string> => {
    try {
      setIsStreaming(true);
      setError(null);

      const settings = loadSettings();

      // Resolve provider + key from the active custom model (if any)
      const activeCustomModel = (settings.customModels || []).find(
        m => m.id === settings.activeCustomModelId
      );
      const apiKey = activeCustomModel?.apiKey || settings.apiKeys[settings.activeProvider] || '';
      const modelId = activeCustomModel?.modelId ?? settings.activeModel;

      // Use GenericOpenAIProvider if custom baseUrl is set, else fall back to registered provider
      const provider = activeCustomModel?.baseUrl
        ? new GenericOpenAIProvider(activeCustomModel.baseUrl, apiKey)
        : getProvider(activeCustomModel?.provider ?? settings.activeProvider);

      if (provider.requiresKey() && !apiKey) {
        throw new Error(`No API key — add one in Settings → My Models.`);
      }
      if (apiKey && !activeCustomModel?.baseUrl) provider.setApiKey(apiKey);

      const requestOptions: AIRequestOptions = {
        model: modelId,
        messages: messages.map(m => ({
          role: m.role as 'system' | 'user' | 'assistant',
          content: m.content,
        })),
        temperature: 0.7,
        maxTokens: 2048,
        stream: false,
        ...options,
      };

      const response = await provider.chat(requestOptions);
      setIsStreaming(false);
      return response.content;
    } catch (err) {
      setIsStreaming(false);
      const msg = (err as Error).message;
      setError(msg);
      throw err;
    }
  }, []);

  const sendMessageStream = useCallback(async (
    messages: ChatMessage[],
    options: any,
    onChunk: (chunk: string) => void
  ): Promise<string> => {
    try {
      setIsStreaming(true);
      setError(null);
      setCurrentStreamText("");
      streamTextRef.current = "";

      abortControllerRef.current = new AbortController();

      const settings = loadSettings();

      // Resolve provider + key from the active custom model (if any)
      const activeCustomModel = (settings.customModels || []).find(
        m => m.id === settings.activeCustomModelId
      );
      const apiKey = activeCustomModel?.apiKey || settings.apiKeys[settings.activeProvider] || '';
      const modelId = activeCustomModel?.modelId ?? settings.activeModel;

      // Use GenericOpenAIProvider if custom baseUrl is set, else fall back to registered provider
      const provider = activeCustomModel?.baseUrl
        ? new GenericOpenAIProvider(activeCustomModel.baseUrl, apiKey)
        : getProvider(activeCustomModel?.provider ?? settings.activeProvider);

      if (provider.requiresKey() && !apiKey) {
        throw new Error(`No API key — add one in Settings → My Models.`);
      }
      if (apiKey && !activeCustomModel?.baseUrl) provider.setApiKey(apiKey);

      const requestOptions: AIRequestOptions = {
        model: modelId,
        messages: messages.map(m => ({
          role: m.role as 'system' | 'user' | 'assistant',
          content: m.content,
        })),
        temperature: 0.7,
        maxTokens: 2048,
        stream: true,
        ...options,
      };

      let fullText = "";
      const stream = provider.chatStream(requestOptions);

      for await (const chunk of stream) {
        if (abortControllerRef.current?.signal.aborted) {
          break;
        }
        if (chunk.content) {
          fullText += chunk.content;
          streamTextRef.current = fullText;
          setCurrentStreamText(fullText);
          onChunk(chunk.content);
        }
        if (chunk.done) break;
      }

      setIsStreaming(false);
      abortControllerRef.current = null;
      return fullText;
    } catch (err) {
      setIsStreaming(false);
      abortControllerRef.current = null;
      const msg = (err as Error).message;
      if (msg !== "Stream cancelled") {
        setError(msg);
      }
      return streamTextRef.current;
    }
  }, []);

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

  const testConnection = useCallback(async (provider: AIProviderType, apiKey: string): Promise<{ ok: boolean; error?: string }> => {
    try {
      const p = getProvider(provider);
      p.setApiKey(apiKey);
      // Use a small, widely-available model for the test ping
      const models = p.getModels().filter(m => m.id !== '__custom__');
      const testModel = models.find(m => m.id.includes('8b') || m.id.includes('7b') || m.id.includes('mini'))
        || models[models.length - 2]  // second to last (before __custom__)
        || models[0];
      const response = await p.chat({
        model: testModel?.id || '',
        messages: [{ role: 'user', content: 'Hi' }],
        maxTokens: 5,
        stream: false,
      });
      return { ok: !!response.content };
    } catch (err) {
      return { ok: false, error: (err as Error).message };
    }
  }, []);

  return {
    sendMessage,
    sendMessageStream,
    cancelStream,
    isStreaming,
    currentStreamText,
    error,
    switchProvider,
    switchModel,
    testConnection,
  };
}
