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
          await ExcelService.writeToRange(cmd.cell, [[cmd.value]], cmd.sheet_name);
          results.executed++;
          break;
        case 'write_formula':
          await ExcelService.insertFormula(cmd.cell, cmd.formula, cmd.sheet_name);
          results.executed++;
          break;
        case 'write_range':
          await ExcelService.writeToRange(cmd.range, cmd.values, cmd.sheet_name);
          results.executed++;
          break;
        case 'create_chart':
          await ExcelService.createChart(cmd.chart_type, cmd.data_range, cmd.title, cmd.sheet_name);
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
          await ExcelService.clearRange(cmd.range, cmd.sheet_name);
          results.executed++;
          break;
        case 'format_range':
          await ExcelService.formatRange(cmd.range, cmd.options, cmd.sheet_name);
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
        case 'switch_sheet':
          await ExcelService.switchSheet(cmd.sheet_name);
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
        case 'replace_entire_body':
          await WordService.replaceEntireBody(cmd.text);
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
        case 'edit_slide_text': {
          await PowerPointService.editSlideText(cmd.slide_index ?? 0, cmd.text, cmd.shape_name);
          results.executed++;
          break;
        }
        case 'get_all_slides': {
          const allSlides = await PowerPointService.getAllSlidesText();
          results.executed++;
          results.errors.push(`[All slides]: ${JSON.stringify(allSlides)}`);
          break;
        }
        default:
          results.errors.push(`Unknown PPT action: ${cmd.action}`);
      }
    } catch (e) {
      results.errors.push((e as Error).message);
    }
  }
  return results;
}

/** Strip <*_CMD>...</*_CMD> blocks from displayed text.
 *  Also strips unclosed command blocks (e.g. when the model is cut off
 *  mid-generation) so the raw <EXCEL_CMD> tag doesn't leak into the UI. */
export function cleanResponseText(text: string): string {
  return text
    .replace(/<(EXCEL|WORD|PPT)_CMD>([/\s\S]*?)<\/\1_CMD>/g, '')   // closed blocks
    .replace(/<(EXCEL|WORD|PPT)_CMD>[/\s\S]*$/g, '')               // unclosed blocks (to end)
    .trim();
}

/** Strip <think>...</think> tags from the thinking field for display.
 *  The tags themselves are captured as thinking chunks (see parseOpenAISSEStream)
 *  so we know where the block starts/ends — but the user doesn't need to see
 *  the literal <think> tags in the UI. */
function cleanThinkingText(thinking: string): string {
  return thinking.replace(/<\/?think>/g, '').trim();
}

/** Execute any host-specific command blocks found in `text`.
 *  Returns the result { executed, errors } so the caller can surface
 *  per-command failures to the user. */
async function executeHostCommands(hostApp: OfficeHostType, text: string): Promise<{ executed: number; errors: string[] } | null> {
  try {
    if (hostApp === 'Excel') return await executeExcelCommands(text);
    else if (hostApp === 'Word') return await executeWordCommands(text);
    else if (hostApp === 'PowerPoint') return await executePPTCommands(text);
  } catch (e) {
    console.warn('Host command execution failed:', e);
  }
  return null;
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
  // True during the post-LLM finishing phase: command execution + state patch.
  // Combined with ai.isStreaming, this gives a complete "is busy" signal.
  const [isFinishing, setIsFinishing] = useState(false);

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

  // Debounce localStorage writes during streaming so we don't do a
  // JSON.stringify + setItem on every single chunk (which can be 1000+
  // times for a long response). The ref + React state still update
  // synchronously so the UI stays live — only the disk write is debounced.
  const storageWriteTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const flushStorageWrite = useCallback(() => {
    if (storageWriteTimer.current) {
      clearTimeout(storageWriteTimer.current);
      storageWriteTimer.current = null;
    }
    saveConversations(conversationsRef.current);
  }, []);

  useEffect(() => {
    const loaded = loadConversations();
    setConversations(loaded);
    if (loaded.length > 0) {
      setActiveConversationId(loaded[0].id);
    }
    // Flush any pending debounced storage write on unmount so we don't
    // lose the last few streamed chunks if the user closes the taskpane.
    return () => {
      if (storageWriteTimer.current) {
        clearTimeout(storageWriteTimer.current);
        saveConversations(conversationsRef.current);
      }
    };
  }, []);

  const activeConversation = conversations.find(c => c.id === activeConversationId) || null;
  const messages = activeConversation?.messages || [];

  /** Persist conversations, enforcing the max-history cap from settings.
   *  Updates conversationsRef.current SYNCHRONOUSLY so that subsequent
   *  reads in the same async tick (e.g. inside sendChatMessage) see the
   *  latest state. The localStorage write is debounced to avoid a write
   *  storm during streaming. */
  const persistConversations = useCallback((next: ChatConversation[]) => {
    const cap = settingsRef.current.maxConversationHistory || 50;
    const trimmed = next.length > cap ? next.slice(0, cap) : next;
    conversationsRef.current = trimmed;
    setConversations(trimmed);
    // Skip localStorage writes if the user disabled conversation persistence.
    if (settingsRef.current.saveConversations === false) return;
    // Debounce the localStorage write — coalesce rapid streaming updates
    // into a single write 300ms after the last chunk.
    if (storageWriteTimer.current) clearTimeout(storageWriteTimer.current);
    storageWriteTimer.current = setTimeout(() => {
      saveConversations(conversationsRef.current);
      storageWriteTimer.current = null;
    }, 300);
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
    //    Apply conversation compaction: if the recent context is too long,
    //    ask the AI to summarize older messages into a compaction_summary
    //    marker so we don't blow the context window.
    const conv = conversationsRef.current.find(c => c.id === convId);
    let historyMessages: ChatMessage[] = conv?.messages ?? [];

    // Drop the empty assistant placeholder we just added — it should not be
    // sent to the LLM as context. BUT: don't patch state with this filtered
    // array, because that would remove the placeholder from the UI. The
    // placeholder must stay in state so finishAssistant can update it later.
    const llmHistory = historyMessages.filter(m => m.id !== assistantMsgId);

    // If compaction is needed, run it (best-effort; if it fails we proceed
    // with the full history and let the provider's own token limit kick in).
    try {
      const { shouldCompact, buildCompactionPrompt, insertCompactionSummary, sliceContextForLLM } =
        await import('../services/compaction');
      const cutIdx = shouldCompact(llmHistory);
      if (cutIdx !== null) {
        const summarizable = llmHistory.slice(0, cutIdx);
        const summaryPrompt = buildCompactionPrompt(summarizable);
        const summaryResp = await ai.sendMessage(
          [
            { id: 'system', role: 'system', content: 'You are a conversation summarizer.', timestamp: 0 },
            { id: 'user', role: 'user', content: summaryPrompt, timestamp: 0 },
          ],
          { maxTokens: 600 },
        );
        const compactedHistory = insertCompactionSummary(llmHistory, cutIdx, summaryResp);
        // Persist the compaction marker INTO the real conversation state —
        // but KEEP the assistant placeholder. Insert the marker before the
        // placeholder's position in the real messages array.
        patchConversation(convId!, c => {
          const realMessages = c.messages;
          const placeholderIdx = realMessages.findIndex(m => m.id === assistantMsgId);
          // Build the new messages array: everything from compactedHistory
          // (which has the compaction marker) + the assistant placeholder.
          const beforePlaceholder = compactedHistory;
          const placeholder = placeholderIdx >= 0 ? realMessages[placeholderIdx] : assistantMessage;
          return { ...c, messages: [...beforePlaceholder, placeholder] };
        });
        // Use the compacted history (without placeholder) for the LLM request.
        historyMessages = compactedHistory;
      } else {
        historyMessages = llmHistory;
      }
      // Slice to LLM-visible context (everything from last compaction onward).
      const llmVisible = sliceContextForLLM(historyMessages);
      const aiMessages: ChatMessage[] = [
        { id: 'system', role: 'system', content: systemPrompt, timestamp: 0 },
        ...llmVisible,
      ];

      // 5. Stream or fetch, then execute commands + clean the display text.
      const finishAssistant = async (raw: string, thinking: string = '') => {
        setIsFinishing(true);
        try {
          const cmdResult = await executeHostCommands(hostApp, raw);
          const stripped = raw.replace(/<think>[\s\S]*?<\/think>/g, '').trim();
          const displayText = cleanResponseText(stripped);
          // Surface command execution results so the user knows the AI actually
          // did something (not just chatted). Show success count AND errors.
          let commandNote = '';
          if (cmdResult && cmdResult.executed > 0) {
            if (cmdResult.errors.length > 0) {
              commandNote = `\n\n---\n✅ ${cmdResult.executed} command(s) executed. ⚠️ ${cmdResult.errors.length} error(s): ${cmdResult.errors.slice(0, 3).join('; ')}${cmdResult.errors.length > 3 ? '...' : ''}`;
            } else {
              commandNote = `\n\n---\n✅ ${cmdResult.executed} command(s) executed successfully.`;
            }
          }
          const cleanThinking = thinking ? cleanThinkingText(thinking) : '';
          patchConversation(convId!, c => ({
            ...c,
            messages: c.messages.map(m =>
              m.id === assistantMsgId
                ? { ...m, content: displayText + commandNote, thinking: cleanThinking || undefined }
                : m,
            ),
          }));
          flushStorageWrite();
        } finally {
          setIsFinishing(false);
          ai.markProcessingDone();
        }
      };

      try {
        if (activeSettings.streamResponses) {
          const onChunk = (chunk: string) => {
            // Use persistConversations (not raw setConversations) so the ref
            // stays in sync — otherwise finishAssistant's patchConversation
            // would read a stale ref and lose the streamed content.
            const next = conversationsRef.current.map(c => {
              if (c.id !== convId) return c;
              return {
                ...c,
                messages: c.messages.map(m =>
                  m.id === assistantMsgId
                    ? { ...m, content: m.content + chunk }
                    : m,
                ),
              };
            });
            persistConversations(next);
          };
          const onThinking = (chunk: string) => {
            const next = conversationsRef.current.map(c => {
              if (c.id !== convId) return c;
              return {
                ...c,
                messages: c.messages.map(m =>
                  m.id === assistantMsgId
                    ? { ...m, thinking: (m.thinking || '') + chunk }
                    : m,
                ),
              };
            });
            persistConversations(next);
          };
          const { text, thinking } = await ai.sendMessageStream(
            aiMessages,
            { webSearch: webSearchEnabled },
            onChunk,
            onThinking,
          );
          await finishAssistant(text, thinking);
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
    } catch (err) {
      // Compaction failed — fall back to the legacy behavior of sending
      // the full history. Better to risk a context overflow than to block
      // the user from sending their message.
      console.warn('Compaction skipped:', err);
      const aiMessages: ChatMessage[] = [
        { id: 'system', role: 'system', content: systemPrompt, timestamp: 0 },
        ...llmHistory,
      ];
      setIsFinishing(true);
      try {
        const full = activeSettings.streamResponses
          ? (await ai.sendMessageStream(aiMessages, { webSearch: webSearchEnabled }, () => {})).text
          : await ai.sendMessage(aiMessages, { webSearch: webSearchEnabled });
        const cmdResult = await executeHostCommands(hostApp, full);
        const displayText = cleanResponseText(full.replace(/<think>[\s\S]*?<\/think>/g, '').trim());
        let commandNote = '';
        if (cmdResult && cmdResult.executed > 0) {
          commandNote = cmdResult.errors.length > 0
            ? `\n\n---\n✅ ${cmdResult.executed} command(s) executed. ⚠️ ${cmdResult.errors.length} error(s): ${cmdResult.errors.slice(0, 3).join('; ')}`
            : `\n\n---\n✅ ${cmdResult.executed} command(s) executed successfully.`;
        }
        patchConversation(convId!, c => ({
          ...c,
          messages: c.messages.map(m =>
            m.id === assistantMsgId ? { ...m, content: displayText + commandNote } : m,
          ),
        }));
        flushStorageWrite();
      } catch (err2) {
        patchConversation(convId!, c => ({
          ...c,
          messages: c.messages.map(m =>
            m.id === assistantMsgId
              ? { ...m, content: `⚠️ Error: ${(err2 as Error).message}` }
              : m,
          ),
        }));
        flushStorageWrite();
      } finally {
        setIsFinishing(false);
        ai.markProcessingDone();
      }
    }
  }, [hostApp, ai, persistConversations, patchConversation, flushStorageWrite]);

  const deleteConversation = useCallback((id: string) => {
    const updated = conversationsRef.current.filter(c => c.id !== id);
    persistConversations(updated);
    flushStorageWrite();
    if (activeIdRef.current === id) {
      setActiveConversationId(updated.length > 0 ? updated[0].id : null);
    }
  }, [persistConversations, flushStorageWrite]);

  const clearAllConversations = useCallback(() => {
    persistConversations([]);
    flushStorageWrite();
    setActiveConversationId(null);
  }, [persistConversations, flushStorageWrite]);

  const exportConversation = useCallback((id: string, format: 'txt' | 'json' | 'markdown') => {
    const conv = conversationsRef.current.find(c => c.id === id);
    if (!conv) return;

    let dataStr = "";
    let ext = format;
    if (format === 'json') {
      dataStr = JSON.stringify(conv, null, 2);
    } else if (format === 'markdown') {
      dataStr = conv.messages
        .filter(m => m.role !== 'system' && m.role !== 'compaction_summary')
        .map(m => `## ${m.role === 'user' ? 'You' : 'AI'}\n\n${m.content}`)
        .join('\n\n---\n\n');
      ext = 'md' as any;
    } else {
      dataStr = conv.messages
        .filter(m => m.role !== 'system' && m.role !== 'compaction_summary')
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
    // Expose isStreaming + cancelStream from the INTERNAL useAI() instance
    // so ChatPanel reads the SAME state that sendChatMessage updates.
    // Without this, ChatPanel would call useAI() separately and get its own
    // independent isStreaming=false that never changes — making it impossible
    // for the user to tell when the LLM is responding vs. done.
    isStreaming: ai.isStreaming,
    cancelStream: ai.cancelStream,
    // isFinishing = true during post-LLM finishing (command execution + patch)
    // isBusy = true for the entire response lifecycle (streaming + finishing)
    isFinishing,
    isBusy: ai.isStreaming || isFinishing,
  };
}

export type UseChatReturn = ReturnType<typeof useChat>;
