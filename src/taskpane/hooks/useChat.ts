import { useState, useCallback, useEffect } from 'react';
import { ChatMessage, ChatConversation, OfficeHostType } from '../types';
import { loadConversations, saveConversations, generateId, loadSettings } from '../utils/storage';
import { useAI } from './useAI';
import { ExcelService } from '../services/office/ExcelService';

// ---- Excel Command Execution ----

export interface ExcelCmdResult { executed: number; errors: string[]; }

export async function executeExcelCommands(text: string): Promise<ExcelCmdResult> {
  const results: ExcelCmdResult = { executed: 0, errors: [] };
  const cmdRegex = /<EXCEL_CMD>([\/\s\S]*?)<\/EXCEL_CMD>/g;
  let match;
  while ((match = cmdRegex.exec(text)) !== null) {
    try {
      const cmd = JSON.parse(match[1].trim());
      switch (cmd.action) {
        case 'write_cell':
          await ExcelService.writeToRange(cmd.cell, [[cmd.value]]);
          results.executed++;
          break;
        case 'write_formula':
          await ExcelService.insertFormula(cmd.cell, cmd.formula);
          results.executed++;
          break;
        case 'write_range':
          await ExcelService.writeToRange(cmd.range, cmd.values);
          results.executed++;
          break;
        case 'create_chart':
          await ExcelService.createChart(cmd.chart_type, cmd.data_range, cmd.title);
          results.executed++;
          break;
        case 'create_pivot_table':
          await ExcelService.createPivotTable(cmd.source_range, cmd.target_cell, cmd.row_field, cmd.value_field, cmd.pivot_name);
          results.executed++;
          break;
        case 'clear_range':
          await ExcelService.clearRange(cmd.range);
          results.executed++;
          break;
        case 'format_range':
          await ExcelService.formatRange(cmd.range, cmd.options);
          results.executed++;
          break;
        case 'add_sheet':
          await ExcelService.addSheet(cmd.name);
          results.executed++;
          break;
        case 'delete_sheet':
          await ExcelService.deleteSheet(cmd.name);
          results.executed++;
          break;
        case 'insert_range':
          await ExcelService.insertRange(cmd.range, cmd.shift_direction);
          results.executed++;
          break;
        case 'delete_range':
          await ExcelService.deleteRange(cmd.range, cmd.shift_direction);
          results.executed++;
          break;
        case 'merge_cells':
          await ExcelService.mergeCells(cmd.range, cmd.merge_across);
          results.executed++;
          break;
        case 'create_table':
          await ExcelService.createTable(cmd.range, cmd.has_headers, cmd.name);
          results.executed++;
          break;
        case 'sort_range':
          await ExcelService.sortRange(cmd.range, cmd.column_index, cmd.ascending);
          results.executed++;
          break;
        case 'find_replace':
          await ExcelService.findAndReplace(cmd.range, cmd.find_text, cmd.replace_text);
          results.executed++;
          break;
        case 'add_data_validation':
          await ExcelService.addDataValidation(cmd.range, cmd.source_list);
          results.executed++;
          break;
        case 'add_conditional_formatting':
          await ExcelService.addConditionalFormatting(cmd.range, cmd.type);
          results.executed++;
          break;
        default:
          results.errors.push(`Unknown action: ${cmd.action}`);
      }
    } catch (e) {
      results.errors.push((e as Error).message);
    }
  }
  return results;
}

/** Strip <EXCEL_CMD>...</EXCEL_CMD> blocks from displayed text */
export function cleanResponseText(text: string): string {
  return text.replace(/<EXCEL_CMD>[\/\s\S]*?<\/EXCEL_CMD>/g, '').trim();
}

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

    // Build system prompt — tell the AI it can directly modify Excel via EXCEL_CMD blocks
    const excelCommandDocs = hostApp === 'Excel' ? `

You can directly modify the user's Excel spreadsheet by emitting EXCEL_CMD blocks in your response.
When the user asks you to write, type, insert, or change data in a cell or range, ALWAYS emit an EXCEL_CMD block to do it automatically.

Available actions:
- Write a value:   <EXCEL_CMD>{"action":"write_cell","cell":"G4","value":"Hello World"}</EXCEL_CMD>
- Write a formula: <EXCEL_CMD>{"action":"write_formula","cell":"A1","formula":"=SUM(B1:B10)"}</EXCEL_CMD>
- Write a range:   <EXCEL_CMD>{"action":"write_range","range":"A1:C3","values":[[1,2,3],[4,5,6],[7,8,9]]}</EXCEL_CMD>
- Create a chart:  <EXCEL_CMD>{"action":"create_chart","chart_type":"column","data_range":"A1:B10","title":"My Chart"}</EXCEL_CMD>
    - Supported chart types: "column", "pie", "line", "bar", "area", "scatter"
- Create a PivotTable: <EXCEL_CMD>{"action":"create_pivot_table","source_range":"A1:D100","target_cell":"F1","row_field":"Category","value_field":"Sales","pivot_name":"SalesSummary"}</EXCEL_CMD>
- Clear a range:   <EXCEL_CMD>{"action":"clear_range","range":"A1:Z100"}</EXCEL_CMD>
- Format a range:  <EXCEL_CMD>{"action":"format_range","range":"A1:A10","options":{"bold":true,"backgroundColor":"#FFFF00","fontColor":"#FF0000","fontSize":14,"wrapText":true,"horizontalAlignment":"Center","numberFormat":"$#,##0.00"}}</EXCEL_CMD>
- Add Sheet:       <EXCEL_CMD>{"action":"add_sheet","name":"NewData"}</EXCEL_CMD>
- Delete Sheet:    <EXCEL_CMD>{"action":"delete_sheet","name":"OldData"}</EXCEL_CMD>
- Insert Range:    <EXCEL_CMD>{"action":"insert_range","range":"A1:A10","shift_direction":"Down"}</EXCEL_CMD>
- Delete Range:    <EXCEL_CMD>{"action":"delete_range","range":"B1:B10","shift_direction":"Left"}</EXCEL_CMD>
- Merge Cells:     <EXCEL_CMD>{"action":"merge_cells","range":"A1:D1","merge_across":false}</EXCEL_CMD>
- Create Table:    <EXCEL_CMD>{"action":"create_table","range":"A1:D100","has_headers":true,"name":"SalesTable"}</EXCEL_CMD>
- Sort Range:      <EXCEL_CMD>{"action":"sort_range","range":"A2:D100","column_index":0,"ascending":true}</EXCEL_CMD>
- Find & Replace:  <EXCEL_CMD>{"action":"find_replace","range":"A1:Z100","find_text":"USA","replace_text":"United States"}</EXCEL_CMD>
- Data Validation: <EXCEL_CMD>{"action":"add_data_validation","range":"B2:B100","source_list":"Yes,No,Maybe"}</EXCEL_CMD>
- Cond. Format:    <EXCEL_CMD>{"action":"add_conditional_formatting","range":"C2:C100","type":"colorScale"}</EXCEL_CMD>

Rules:
- ALWAYS emit an EXCEL_CMD block when the user asks you to put/type/write/insert/set data in Excel.
- You can emit multiple EXCEL_CMD blocks in one response.
- After each block, briefly confirm what you did (e.g. "I've written 'Hello World' to cell G4.").
- If the cell address is ambiguous, use the most likely one based on context.
- Never ask the user to do it manually if you can do it with an EXCEL_CMD block.` : '';

    const systemPrompt = `You are an Your Co-Pilot assistant for Microsoft ${hostApp}. You help users with data analysis, formulas, writing, presentations, and more. Be helpful, concise, and provide actionable answers. When providing code, formulas, or structured data, use markdown formatting.${excelCommandDocs}${contextStr ? `\n\nCurrent document context:\n${contextStr}` : ''}`;

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

        // Final save — execute Excel commands then clean response
        if (hostApp === 'Excel') {
          try { await executeExcelCommands(fullResponse); } catch {}
        }
        const displayText = cleanResponseText(fullResponse);
        const finalConvs = currentConvs.map(c => {
          if (c.id === convId) {
            return {
              ...c,
              messages: c.messages.map(m =>
                m.id === assistantMsgId ? { ...m, content: displayText } : m
              ),
            };
          }
          return c;
        });
        saveConversations(finalConvs);
      } else {
        // Non-streaming mode — execute Excel commands then clean response
        const response = await ai.sendMessage(aiMessages);
        if (hostApp === 'Excel') {
          try { await executeExcelCommands(response); } catch {}
        }
        const displayText = cleanResponseText(response);
        currentConvs = currentConvs.map(c => {
          if (c.id === convId) {
            return {
              ...c,
              messages: c.messages.map(m =>
                m.id === assistantMsgId ? { ...m, content: displayText } : m
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
