import { useState, useCallback, useEffect, useRef } from 'react';
import { ChatMessage, ChatConversation, OfficeHostType } from '../types';
import { useSettings } from './useSettings';
import { loadConversations, saveConversations, generateId } from '../utils/storage';
import { useAI } from './useAI';
import { buildSystemPrompt } from '../services/prompts';
import { ExcelService } from '../services/office/ExcelService';

// ---- Excel Command Execution ----

export interface ExcelCmdResult { executed: number; errors: string[]; }

export async function executeExcelCommands(text: string): Promise<ExcelCmdResult> {
  const results: ExcelCmdResult = { executed: 0, errors: [] };
  const cmdRegex = /<EXCEL_CMD>([/\s\S]*?)<\/EXCEL_CMD>/g;
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
        case 'delete_chart':
          await ExcelService.deleteChart(cmd.chart_name);
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
        case 'remove_duplicates':
          await ExcelService.removeDuplicates(cmd.range, cmd.columns);
          results.executed++;
          break;
        case 'trim_whitespace':
          await ExcelService.trimWhitespace(cmd.range);
          results.executed++;
          break;
        case 'change_case':
          await ExcelService.changeCase(cmd.range, cmd.type);
          results.executed++;
          break;
        case 'remove_blank_rows':
          await ExcelService.removeBlankRows(cmd.range);
          results.executed++;
          break;
        case 'apply_filter':
          await ExcelService.applyFilter(cmd.range, cmd.column_index, cmd.criteria);
          results.executed++;
          break;
        case 'clear_filter':
          await ExcelService.clearFilter();
          results.executed++;
          break;
        case 'group_data':
          await ExcelService.groupData(cmd.range, cmd.by_rows);
          results.executed++;
          break;
        case 'ungroup_data':
          await ExcelService.ungroupData(cmd.range, cmd.by_rows);
          results.executed++;
          break;
        case 'add_sparklines':
          await ExcelService.addSparklines(cmd.range, cmd.source_range, cmd.type);
          results.executed++;
          break;
        case 'format_chart':
          await ExcelService.formatChart(cmd.chart_name, cmd.options);
          results.executed++;
          break;
        case 'highlight_duplicates':
          await ExcelService.highlightDuplicates(cmd.range, cmd.color);
          results.executed++;
          break;
        case 'highlight_top_bottom':
          await ExcelService.highlightTopBottom(cmd.range, cmd.type, cmd.count, cmd.color);
          results.executed++;
          break;
        // ── From office-agents ──────────────────────────────────────────────
        case 'eval_js': {
          const jsResult = await ExcelService.evalOfficeJs(cmd.code);
          results.executed++;
          if (jsResult !== null && jsResult !== undefined) {
            results.errors.push(`[eval_js result]: ${JSON.stringify(jsResult)}`);
          }
          break;
        }
        case 'screenshot_range': {
          // Screenshot is handled in the response rendering — just mark executed
          await ExcelService.screenshotRange(cmd.range);
          results.executed++;
          break;
        }
        case 'search_data': {
          const searchResults = await ExcelService.searchData(cmd.search_term, {
            matchCase: cmd.match_case,
            useRegex: cmd.use_regex,
            range: cmd.range,
            maxResults: cmd.max_results,
          });
          results.executed++;
          results.errors.push(`[search result]: ${JSON.stringify(searchResults)}`);
          break;
        }
        case 'get_all_objects': {
          const objects = await ExcelService.getAllObjects();
          results.executed++;
          results.errors.push(`[objects]: ${JSON.stringify(objects)}`);
          break;
        }
        case 'get_range_csv': {
          const csv = await ExcelService.getRangeAsCsv(cmd.range, cmd.max_rows);
          results.executed++;
          results.errors.push(`[CSV data]:\n${csv}`);
          break;
        }
        case 'freeze_panes':
          await ExcelService.freezePanes(cmd.cell);
          results.executed++;
          break;
        case 'unfreeze_panes':
          await ExcelService.unfreezePanes();
          results.executed++;
          break;
        case 'autofit_columns':
          await ExcelService.autoFitColumns(cmd.range);
          results.executed++;
          break;
        case 'autofit_rows':
          await ExcelService.autoFitRows(cmd.range);
          results.executed++;
          break;
        case 'create_named_range':
          await ExcelService.upsertNamedRange(cmd.name, cmd.address);
          results.executed++;
          break;
        case 'delete_named_range':
          await ExcelService.deleteNamedRange(cmd.name);
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

// ---- Word Command Execution ----
export async function executeWordCommands(text: string): Promise<ExcelCmdResult> {
  const results: ExcelCmdResult = { executed: 0, errors: [] };
  const cmdRegex = /<WORD_CMD>([/\s\S]*?)<\/WORD_CMD>/g;
  let match;
  while ((match = cmdRegex.exec(text)) !== null) {
    try {
      const cmd = JSON.parse(match[1].trim());
      const { WordService } = await import('../services/office/WordService');
      switch (cmd.action) {
        case 'insert_table':
          await WordService.insertTable(cmd.values, cmd.location || 'end');
          results.executed++;
          break;
        case 'insert_paragraph':
          await WordService.insertParagraph(cmd.text, cmd.location || 'after');
          results.executed++;
          break;
        case 'format_text':
          await WordService.formatText(cmd.options);
          results.executed++;
          break;
        case 'apply_style':
          await WordService.applyStyle(cmd.style);
          results.executed++;
          break;
        case 'clear_formatting':
          await WordService.clearFormatting();
          results.executed++;
          break;
        case 'search_replace':
          await WordService.searchReplace(cmd.find_text, cmd.replace_text);
          results.executed++;
          break;
        case 'eval_js': {
          const jsResult = await WordService.evalOfficeJs(cmd.code);
          results.executed++;
          if (jsResult !== null && jsResult !== undefined) {
            results.errors.push(`[Word eval_js result]: ${JSON.stringify(jsResult)}`);
          }
          break;
        }
        case 'get_structure': {
          const struct = await WordService.getDocumentStructure();
          results.executed++;
          results.errors.push(`[Word structure]: ${JSON.stringify(struct)}`);
          break;
        }
        case 'highlight_search': {
          const count = await WordService.searchAndHighlight(cmd.find_text, cmd.color || 'Yellow');
          results.executed++;
          results.errors.push(`[Highlighted ${count} matches]`);
          break;
        }
        default:
          results.errors.push(`Unknown Word action: ${cmd.action}`);
      }
    } catch (e) {
      results.errors.push((e as Error).message);
    }
  }
  return results;
}

// ---- PPT Command Execution ----
export async function executePPTCommands(text: string): Promise<ExcelCmdResult> {
  const results: ExcelCmdResult = { executed: 0, errors: [] };
  const cmdRegex = /<PPT_CMD>([/\s\S]*?)<\/PPT_CMD>/g;
  let match;
  while ((match = cmdRegex.exec(text)) !== null) {
    try {
      const cmd = JSON.parse(match[1].trim());
      const { PowerPointService } = await import('../services/office/PowerPointService');
      switch (cmd.action) {
        case 'add_slide':
          await PowerPointService.addSlide();
          results.executed++;
          break;
        case 'add_textbox':
          await PowerPointService.addTextbox(cmd.text);
          results.executed++;
          break;
        case 'add_shape':
          await PowerPointService.addShape(cmd.shape_type);
          results.executed++;
          break;
        case 'format_shape':
          await PowerPointService.formatShape(cmd.shape_index, cmd.fill_color, cmd.font_color);
          results.executed++;
          break;
        case 'set_slide_notes':
          await PowerPointService.setSlideNotes(cmd.notes);
          results.executed++;
          break;
        case 'eval_js': {
          const jsResult = await PowerPointService.evalOfficeJs(cmd.code);
          results.executed++;
          if (jsResult !== null && jsResult !== undefined) {
            results.errors.push(`[PPT eval_js result]: ${JSON.stringify(jsResult)}`);
          }
          break;
        }
        case 'get_shapes': {
          const shapes = await PowerPointService.getAllSlideShapes(cmd.slide_index || 0);
          results.executed++;
          results.errors.push(`[Slide shapes]: ${JSON.stringify(shapes)}`);
          break;
        }
        case 'delete_slide':
          await PowerPointService.deleteSlide(cmd.slide_index || 0);
          results.executed++;
          break;
        default:
          results.errors.push(`Unknown PPT action: ${cmd.action}`);
      }
    } catch (e) {
      results.errors.push((e as Error).message);
    }
  }
  return results;
}

/** Strip <*_CMD>...</*_CMD> blocks from displayed text */
export function cleanResponseText(text: string): string {
  return text.replace(/<(EXCEL|WORD|PPT)_CMD>([/\s\S]*?)<\/\1_CMD>/g, '').trim();
}

/** Execute any host-specific command blocks found in `text`. */
async function executeHostCommands(hostApp: OfficeHostType, text: string): Promise<void> {
  try {
    if (hostApp === 'Excel') await executeExcelCommands(text);
    else if (hostApp === 'Word') await executeWordCommands(text);
    else if (hostApp === 'PowerPoint') await executePPTCommands(text);
  } catch (e) {
    // Command execution errors are surfaced via the results.errors array,
    // not thrown to the caller — the chat response should still display.
    console.warn('Host command execution failed:', e);
  }
}

/** Fetch document context (selected range / body / slides) for the host. */
async function fetchHostContext(hostApp: OfficeHostType): Promise<string> {
  if (hostApp === 'Excel') {
    const { ExcelService } = await import('../services/office/ExcelService');
    return await ExcelService.getContextForAI();
  }
  if (hostApp === 'Word') {
    const { WordService } = await import('../services/office/WordService');
    return await WordService.getContextForAI();
  }
  if (hostApp === 'PowerPoint') {
    const { PowerPointService } = await import('../services/office/PowerPointService');
    return await PowerPointService.getContextForAI();
  }
  return '';
}

export function useChat(hostApp: OfficeHostType) {
  const [conversations, setConversations] = useState<ChatConversation[]>([]);
  const [activeConversationId, setActiveConversationId] = useState<string | null>(null);
  const [includeContext, setIncludeContext] = useState(true);

  const ai = useAI();
  const { settings } = useSettings();

  // Keep latest settings/conversations accessible inside async callbacks
  // without re-creating the callback on every state change.
  const settingsRef = useRef(settings);
  settingsRef.current = settings;
  const conversationsRef = useRef(conversations);
  conversationsRef.current = conversations;
  const activeIdRef = useRef(activeConversationId);
  activeIdRef.current = activeConversationId;

  useEffect(() => {
    const loaded = loadConversations();
    setConversations(loaded);
    if (loaded.length > 0) {
      setActiveConversationId(loaded[0].id);
    }
  }, []);

  const activeConversation = conversations.find(c => c.id === activeConversationId) || null;
  const messages = activeConversation?.messages || [];

  /** Persist conversations, enforcing the max-history cap from settings. */
  const persistConversations = useCallback((next: ChatConversation[]) => {
    const cap = settingsRef.current.maxConversationHistory || 50;
    const trimmed = next.length > cap ? next.slice(0, cap) : next;
    setConversations(trimmed);
    saveConversations(trimmed);
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
    persistConversations([newConv, ...conversationsRef.current]);
    setActiveConversationId(newConv.id);
    return newConv;
  }, [hostApp, persistConversations]);

  /** Patch a single conversation by id using an updater fn. */
  const patchConversation = useCallback(
    (id: string, updater: (c: ChatConversation) => ChatConversation) => {
      const next = conversationsRef.current.map(c => (c.id === id ? updater(c) : c));
      persistConversations(next);
    },
    [persistConversations],
  );

  const sendChatMessage = useCallback(async (
    content: string,
    includeContext: boolean = false,
    webSearchEnabled: boolean = false,
  ) => {
    // 1. Gather document context if requested.
    let contextStr = '';
    if (includeContext) {
      try {
        contextStr = await fetchHostContext(hostApp);
      } catch (err) {
        console.warn('Failed to get Office context:', err);
      }
    }

    const systemPrompt = buildSystemPrompt(hostApp, contextStr);
    const activeSettings = settingsRef.current;

    // 2. Ensure an active conversation exists.
    let convId = activeIdRef.current;
    if (!convId) {
      const newConv: ChatConversation = {
        id: generateId(),
        title: content.substring(0, 40),
        messages: [],
        createdAt: Date.now(),
        updatedAt: Date.now(),
        hostApp,
        model: activeSettings.activeModel,
        provider: activeSettings.activeProvider,
      };
      persistConversations([newConv, ...conversationsRef.current]);
      convId = newConv.id;
      setActiveConversationId(convId);
    }

    // 3. Append the user message + an empty assistant placeholder.
    const userMessage: ChatMessage = {
      id: generateId(),
      role: 'user',
      content,
      timestamp: Date.now(),
      contextIncluded: includeContext && !!contextStr,
    };
    const assistantMessage: ChatMessage = {
      id: generateId(),
      role: 'assistant',
      content: '',
      timestamp: Date.now(),
      model: activeSettings.activeModel,
      provider: activeSettings.activeProvider,
    };
    const assistantMsgId = assistantMessage.id;

    patchConversation(convId, c => ({
      ...c,
      messages: [...c.messages, userMessage, assistantMessage],
      updatedAt: Date.now(),
      title: c.messages.length === 0 ? content.substring(0, 40) : c.title,
    }));

    // 4. Build the AI request payload (system + history).
    const conv = conversationsRef.current.find(c => c.id === convId);
    const historyMessages: ChatMessage[] = conv?.messages ?? [];
    const aiMessages: ChatMessage[] = [
      { id: 'system', role: 'system', content: systemPrompt, timestamp: 0 },
      ...historyMessages,
    ];

    // 5. Stream or fetch, then execute commands + clean the display text.
    const finishAssistant = async (raw: string) => {
      await executeHostCommands(hostApp, raw);
      const displayText = cleanResponseText(raw);
      patchConversation(convId!, c => ({
        ...c,
        messages: c.messages.map(m =>
          m.id === assistantMsgId ? { ...m, content: displayText } : m,
        ),
      }));
    };

    try {
      if (activeSettings.streamResponses) {
        const onChunk = (chunk: string) => {
          // Live-update the assistant placeholder using a functional state
          // update so we don't depend on stale `conversations`.
          setConversations(prev => prev.map(c => {
            if (c.id !== convId) return c;
            return {
              ...c,
              messages: c.messages.map(m =>
                m.id === assistantMsgId
                  ? { ...m, content: m.content + chunk }
                  : m,
              ),
            };
          }));
        };
        const full = await ai.sendMessageStream(aiMessages, { webSearch: webSearchEnabled }, onChunk);
        await finishAssistant(full);
      } else {
        const full = await ai.sendMessage(aiMessages, { webSearch: webSearchEnabled });
        await finishAssistant(full);
      }
    } catch (err) {
      patchConversation(convId!, c => ({
        ...c,
        messages: c.messages.map(m =>
          m.id === assistantMsgId
            ? { ...m, content: `⚠️ Error: ${(err as Error).message}` }
            : m,
        ),
      }));
    }
  }, [hostApp, ai, persistConversations, patchConversation]);

  const deleteConversation = useCallback((id: string) => {
    const updated = conversationsRef.current.filter(c => c.id !== id);
    persistConversations(updated);
    if (activeIdRef.current === id) {
      setActiveConversationId(updated.length > 0 ? updated[0].id : null);
    }
  }, [persistConversations]);

  const clearAllConversations = useCallback(() => {
    persistConversations([]);
    setActiveConversationId(null);
  }, [persistConversations]);

  const exportConversation = useCallback((id: string, format: 'txt' | 'json' | 'markdown') => {
    const conv = conversationsRef.current.find(c => c.id === id);
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
  }, []);

  const setActiveConversation = useCallback((id: string) => {
    setActiveConversationId(id);
  }, []);

  return {
    conversations,
    activeConversation,
    messages,
    includeContext,
    setIncludeContext,
    createConversation,
    sendChatMessage,
    deleteConversation,
    clearAllConversations,
    exportConversation,
    setActiveConversation,
  };
}

export type UseChatReturn = ReturnType<typeof useChat>;
