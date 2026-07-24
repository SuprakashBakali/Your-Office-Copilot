import { useState, useCallback, useRef } from 'react';
import { AIProviderType, ChatMessage } from '../types';
import { loadSettings } from '../utils/storage';
import { getProvider } from '../services/ai/ProviderFactory';
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
      const provider = getProvider(settings.activeProvider);
      const apiKey = settings.apiKeys[settings.activeProvider];

      if (provider.requiresKey() && !apiKey) {
        throw new Error(`API key not configured for ${settings.activeProvider}. Go to Settings → API Keys.`);
      }

      if (apiKey) {
        provider.setApiKey(apiKey);
      }

      const requestOptions: AIRequestOptions = {
        model: settings.activeModel,
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
      const provider = getProvider(settings.activeProvider);
      const apiKey = settings.apiKeys[settings.activeProvider];

      if (provider.requiresKey() && !apiKey) {
        throw new Error(`API key not configured for ${settings.activeProvider}. Go to Settings → API Keys.`);
      }

      if (apiKey) {
        provider.setApiKey(apiKey);
      }

      const requestOptions: AIRequestOptions = {
        model: settings.activeModel,
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

  const testConnection = useCallback(async (provider: AIProviderType, apiKey: string): Promise<boolean> => {
    try {
      const p = getProvider(provider);
      p.setApiKey(apiKey);
      const response = await p.chat({
        model: p.getModels()[0]?.id || '',
        messages: [{ role: 'user', content: 'Hello, respond with just "OK".' }],
        maxTokens: 10,
        stream: false,
      });
      return !!response.content;
    } catch {
      return false;
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
