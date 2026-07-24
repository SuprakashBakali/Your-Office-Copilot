import { useState, useCallback, useEffect } from 'react';
import { ChatMessage, ChatConversation, OfficeHostType } from '../types';
import { loadConversations, saveConversations, generateId, loadSettings } from '../utils/storage';
import { useAI } from './useAI';

export function useChat(hostApp: OfficeHostType) {
  const [conversations, setConversations] = useState<ChatConversation[]>([]);
  const [activeConversationId, setActiveConversationId] = useState<string | null>(null);

  const ai = useAI();

  useEffect(() => {
    const loaded = loadConversations();
    setConversations(loaded);
    if (loaded.length > 0) {
      setActiveConversationId(loaded[0].id);
    }
  }, []);

  const activeConversation = conversations.find(c => c.id === activeConversationId) || null;
  const messages = activeConversation?.messages || [];

  const updateConversations = useCallback((newConversations: ChatConversation[]) => {
    setConversations(newConversations);
    saveConversations(newConversations);
  }, []);

  const createConversation = useCallback(() => {
    const newConv: ChatConversation = {
      id: generateId(),
      title: 'New Chat',
      messages: [],
      createdAt: Date.now(),
      updatedAt: Date.now(),
      hostApp,
    };
    const updated = [newConv, ...conversations];
    updateConversations(updated);
    setActiveConversationId(newConv.id);
    return newConv;
  }, [conversations, hostApp, updateConversations]);

  const sendChatMessage = useCallback(async (content: string, includeContext: boolean = false) => {
    let contextStr = "";
    if (includeContext) {
      try {
        // Dynamically import to avoid issues when Office.js isn't available
        if (hostApp === 'Excel') {
          const { ExcelService } = await import('../services/office/ExcelService');
          contextStr = await ExcelService.getContextForAI();
        } else if (hostApp === 'Word') {
          const { WordService } = await import('../services/office/WordService');
          contextStr = await WordService.getContextForAI();
        } else if (hostApp === 'PowerPoint') {
          const { PowerPointService } = await import('../services/office/PowerPointService');
          contextStr = await PowerPointService.getContextForAI();
        }
      } catch (err) {
        console.warn('Failed to get Office context:', err);
      }
    }

    const settings = loadSettings();
    const systemPrompt = `You are an AI Copilot assistant for Microsoft ${hostApp}. You help users with data analysis, formulas, writing, presentations, and more. Be helpful, concise, and provide actionable answers. When providing code, formulas, or structured data, use markdown formatting.${contextStr ? `\n\nCurrent document context:\n${contextStr}` : ''}`;

    // Ensure we have an active conversation
    let convId = activeConversationId;
    let currentConvs = [...conversations];

    if (!convId) {
      const newConv: ChatConversation = {
        id: generateId(),
        title: content.substring(0, 40),
        messages: [],
        createdAt: Date.now(),
        updatedAt: Date.now(),
        hostApp,
        model: settings.activeModel,
        provider: settings.activeProvider,
      };
      currentConvs = [newConv, ...currentConvs];
      convId = newConv.id;
      setActiveConversationId(convId);
    }

    // Add user message
    const userMessage: ChatMessage = {
      id: generateId(),
      role: 'user',
      content,
      timestamp: Date.now(),
      contextIncluded: includeContext && !!contextStr,
    };

    currentConvs = currentConvs.map(c => {
      if (c.id === convId) {
        return {
          ...c,
          messages: [...c.messages, userMessage],
          updatedAt: Date.now(),
          title: c.messages.length === 0 ? content.substring(0, 40) : c.title,
        };
      }
      return c;
    });
    updateConversations(currentConvs);

    // Build messages array for AI
    const conv = currentConvs.find(c => c.id === convId);
    const aiMessages: ChatMessage[] = [
      { id: 'system', role: 'system', content: systemPrompt, timestamp: 0 },
      ...(conv?.messages || []),
    ];

    // Add placeholder assistant message
    const assistantMessage: ChatMessage = {
      id: generateId(),
      role: 'assistant',
      content: '',
      timestamp: Date.now(),
      model: settings.activeModel,
      provider: settings.activeProvider,
    };

    const assistantMsgId = assistantMessage.id;
    currentConvs = currentConvs.map(c => {
      if (c.id === convId) {
        return { ...c, messages: [...c.messages, assistantMessage] };
      }
      return c;
    });
    updateConversations(currentConvs);

    try {
      if (settings.streamResponses) {
        // Streaming mode
        let fullResponse = "";
        await ai.sendMessageStream(aiMessages, {}, (chunk) => {
          fullResponse += chunk;
          // Update the assistant message in real-time
          setConversations(prev => prev.map(c => {
            if (c.id === convId) {
              return {
                ...c,
                messages: c.messages.map(m =>
                  m.id === assistantMsgId ? { ...m, content: fullResponse } : m
                ),
              };
            }
            return c;
          }));
        });

        // Final save
        const finalConvs = currentConvs.map(c => {
          if (c.id === convId) {
            return {
              ...c,
              messages: c.messages.map(m =>
                m.id === assistantMsgId ? { ...m, content: fullResponse } : m
              ),
            };
          }
          return c;
        });
        saveConversations(finalConvs);
      } else {
        // Non-streaming mode
        const response = await ai.sendMessage(aiMessages);
        currentConvs = currentConvs.map(c => {
          if (c.id === convId) {
            return {
              ...c,
              messages: c.messages.map(m =>
                m.id === assistantMsgId ? { ...m, content: response } : m
              ),
            };
          }
          return c;
        });
        updateConversations(currentConvs);
      }
    } catch (err) {
      // Update assistant message with error
      currentConvs = currentConvs.map(c => {
        if (c.id === convId) {
          return {
            ...c,
            messages: c.messages.map(m =>
              m.id === assistantMsgId
                ? { ...m, content: `⚠️ Error: ${(err as Error).message}` }
                : m
            ),
          };
        }
        return c;
      });
      updateConversations(currentConvs);
    }
  }, [activeConversationId, conversations, hostApp, ai, updateConversations]);

  const deleteConversation = useCallback((id: string) => {
    const updated = conversations.filter(c => c.id !== id);
    updateConversations(updated);
    if (activeConversationId === id) {
      setActiveConversationId(updated.length > 0 ? updated[0].id : null);
    }
  }, [conversations, activeConversationId, updateConversations]);

  const clearAllConversations = useCallback(() => {
    updateConversations([]);
    setActiveConversationId(null);
  }, [updateConversations]);

  const exportConversation = useCallback((id: string, format: 'txt' | 'json' | 'markdown') => {
    const conv = conversations.find(c => c.id === id);
    if (!conv) return;

    let dataStr = "";
    let ext = format;
    if (format === 'json') {
      dataStr = JSON.stringify(conv, null, 2);
    } else if (format === 'markdown') {
      dataStr = conv.messages
        .filter(m => m.role !== 'system')
        .map(m => `## ${m.role === 'user' ? 'You' : 'AI'}\n\n${m.content}`)
        .join('\n\n---\n\n');
      ext = 'md' as any;
    } else {
      dataStr = conv.messages
        .filter(m => m.role !== 'system')
        .map(m => `${m.role.toUpperCase()}: ${m.content}`)
        .join('\n\n');
    }

    const dataUri = 'data:text/plain;charset=utf-8,' + encodeURIComponent(dataStr);
    const link = document.createElement('a');
    link.setAttribute('href', dataUri);
    link.setAttribute('download', `chat_${conv.title.replace(/\s+/g, '_')}.${ext}`);
    link.click();
  }, [conversations]);

  const setActiveConversation = useCallback((id: string) => {
    setActiveConversationId(id);
  }, []);

  return {
    conversations,
    activeConversation,
    messages,
    createConversation,
    sendChatMessage,
    deleteConversation,
    clearAllConversations,
    exportConversation,
    setActiveConversation,
  };
}
