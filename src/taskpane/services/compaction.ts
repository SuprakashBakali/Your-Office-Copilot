/**
 * Conversation compaction — when a conversation grows too long for the
 * model's context window, summarize the older messages into a single
 * `<compaction_summary>` block so the recent context is preserved.
 *
 * Design ported from office-agents' TODO.md:
 *   - Messages array stays flat. A `compaction_summary` marker message is
 *     inserted at the cut point.
 *   - Older messages are preserved for display but excluded from LLM context.
 *   - Re-compaction is idempotent: the previous summary is already in
 *     context, so the next compaction naturally folds it in.
 *
 * Trigger heuristic: character count of the concatenated message contents.
 * We use ~12k chars as the soft limit (roughly 3-4k tokens, leaving plenty
 * of headroom for the system prompt + new response). This is conservative
 * because we don't know each model's actual context window.
 */
import { ChatMessage } from '../types';

/** The marker role used for compaction summary messages. */
export const COMPACTION_ROLE = 'compaction_summary' as const;

/** Soft limit on total message content before we trigger compaction. */
export const COMPACTION_CHAR_LIMIT = 12000;

/** How many recent messages to always preserve verbatim (never compact). */
export const COMPACTION_KEEP_RECENT = 6;

/**
 * Estimate the total character count of messages that would be sent to the
 * LLM (system + everything after the last compaction marker).
 */
export function estimateContextChars(messages: ChatMessage[]): number {
  // Find the last compaction marker; only count from there onward.
  let startIdx = 0;
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i].role === COMPACTION_ROLE) {
      startIdx = i;
      break;
    }
  }
  let total = 0;
  for (let i = startIdx; i < messages.length; i++) {
    total += messages[i].content.length;
  }
  return total;
}

/**
 * Decide whether compaction should run.
 * Returns the cut index (messages before this index get summarized) or null.
 */
export function shouldCompact(messages: ChatMessage[]): number | null {
  if (messages.length < COMPACTION_KEEP_RECENT * 2) return null;
  if (estimateContextChars(messages) < COMPACTION_CHAR_LIMIT) return null;

  // Cut point: keep the last COMPACTION_KEEP_RECENT messages verbatim.
  // Walk backwards from there, but never compact a compaction_summary.
  const cut = messages.length - COMPACTION_KEEP_RECENT;
  if (cut <= 0) return null;
  // If the cut would land on a compaction_summary, push it forward.
  let adjusted = cut;
  while (adjusted < messages.length && messages[adjusted].role === COMPACTION_ROLE) {
    adjusted++;
  }
  return adjusted < messages.length ? adjusted : null;
}

/**
 * Build the summarization prompt that turns old messages into a compact
 * summary. The AI is asked to preserve facts, decisions, and any pending
 * action items.
 */
export function buildCompactionPrompt(messagesToSummarize: ChatMessage[]): string {
  const transcript = messagesToSummarize
    .filter(m => m.role !== 'system')
    .map(m => {
      // Use a cleaner label for compaction markers so the summarizer LLM
      // understands it's a previous summary, not a new user/assistant turn.
      if (m.role === 'compaction_summary') {
        return `[PREVIOUS SUMMARY]\n${m.content}`;
      }
      return `[${m.role.toUpperCase()}]\n${m.content}`;
    })
    .join('\n\n---\n\n');

  return `You are a conversation summarizer. Below is a transcript of an earlier portion of a chat between a user and an AI Office Copilot assistant. Summarize it into a concise context block that preserves:

1. The user's overall goal / task
2. Key facts about the document/data the user is working with
3. Decisions made and actions already taken (e.g. "wrote formula to A1", "created chart named SalesChart")
4. Any pending action items or follow-ups
5. Important constraints, preferences, or context the user mentioned

Be specific — preserve cell references, file names, function names, and any URLs. Do NOT include pleasantries or meta-commentary. Output a single markdown block, max ~400 words.

TRANSCRIPT:
${transcript}`;
}

/**
 * Insert a compaction_summary message at the cut index, returning a new
 * array. The older messages are preserved (for display) but will be
 * excluded from LLM context by `sliceContextForLLM`.
 */
export function insertCompactionSummary(
  messages: ChatMessage[],
  cutIndex: number,
  summary: string,
): ChatMessage[] {
  const marker: ChatMessage = {
    id: `compaction-${Date.now()}`,
    role: COMPACTION_ROLE,
    content: summary,
    timestamp: Date.now(),
  };
  return [
    ...messages.slice(0, cutIndex),
    marker,
    ...messages.slice(cutIndex),
  ];
}

/**
 * Return the messages that should be sent to the LLM: the system prompt
 * (handled separately by the caller) plus everything from the last
 * compaction_summary onward. Compaction markers themselves are converted
 * into a system-role message so the LLM sees the summary.
 */
export function sliceContextForLLM(messages: ChatMessage[]): ChatMessage[] {
  let startIdx = 0;
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i].role === COMPACTION_ROLE) {
      startIdx = i;
      break;
    }
  }
  const slice = messages.slice(startIdx);
  // Convert any compaction_summary markers in the slice to system messages
  // so the LLM treats them as context, not as a user turn.
  return slice.map(m =>
    m.role === COMPACTION_ROLE
      ? { ...m, role: 'system' as const, content: `<compaction_summary>\n${m.content}\n</compaction_summary>` }
      : m,
  );
}
