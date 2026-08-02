# Bug Fixes — Round 2

This patch fixes all 8 bugs identified in the audit, plus re-applies the
two fixes from the previous round (nav-bar invisible + messages not showing)
that hadn't been pushed yet.

## All fixes applied

### Previously delivered (re-applied on top of `b6f0da7`)

| Bug | File | Fix |
|-----|------|-----|
| Nav bar invisible (PopoverTrigger/Tooltip ref break) | `PromptTemplatesPanel.tsx`, `NavigationTabs.tsx` | Manual Popover with ref'd Button; swapped Tooltip/MenuTrigger order |
| Messages not showing (stale conversationsRef race) | `useChat.ts` | `persistConversations` now updates `conversationsRef.current` synchronously |

### New fixes (8 bugs from audit)

| # | Bug | Severity | File(s) | Fix |
|---|-----|----------|---------|-----|
| 1 | Thinking-block fragmentation pollutes response | 🔴 HIGH | `ai/types.ts` | SSE parser now tracks `inThinkingBlock` state — content between `<think>` and `</think>` tags is routed to the thinking channel even when tags are split across chunks |
| 2 | Anthropic crashes on empty response | 🔴 HIGH | `AnthropicProvider.ts` | `data.content[0].text` → `data.content?.[0]?.text ?? ''` (optional chaining) |
| 3 | localStorage write storm during streaming | 🔴 HIGH | `useChat.ts` | Debounced `saveConversations` — 300ms coalescing window; ref + React state still update synchronously so UI stays live; flush on unmount + on `finishAssistant` |
| 4 | Silent localStorage quota failure | 🟡 MED | `storage.ts`, `AppContext.tsx` | `setItem` now dispatches a `storage-error` CustomEvent; `AppContext` listens and shows a warning toast |
| 5 | Export includes `compaction_summary` as "AI" | 🟡 MED | `useChat.ts` | Export filter now excludes `compaction_summary` role (both markdown and txt formats) |
| 6 | Compaction feeds `[COMPACTION_SUMMARY]` to LLM | 🟡 MED | `compaction.ts` | `buildCompactionPrompt` now labels compaction markers as `[PREVIOUS SUMMARY]` instead of `[COMPACTION_SUMMARY]` |
| 7 | Malformed command blocks leak into display | 🟡 MED | `useChat.ts` | `cleanResponseText` now strips unclosed `<EXCEL_CMD>...` blocks (to end of string) in addition to closed blocks |
| 8 | MessageBubble crashes on null content | 🟢 LOW | `MessageBubble.tsx` | Guarded `message.content` with `|| ''` fallback; all downstream references use the local `content` variable |

## Verification

```
npx tsc --noEmit              → 0 errors
npx eslint src/**/* --max-warnings 0 → 0 errors, 0 warnings
npx webpack --mode production → succeeds, ~985 KB
```

## Files changed (11 total)

```
src/taskpane/services/ai/types.ts              [MOD] Bug #1: thinking-block parser
src/taskpane/services/ai/AnthropicProvider.ts  [MOD] Bug #2: optional chaining
src/taskpane/hooks/useChat.ts                  [MOD] Bug #3, #5, #7 + re-applied nav-bar/messages fixes
src/taskpane/utils/storage.ts                  [MOD] Bug #4: storage-error event
src/taskpane/store/AppContext.tsx              [MOD] Bug #4: listen for storage-error
src/taskpane/services/compaction.ts            [MOD] Bug #6: [PREVIOUS SUMMARY] label
src/taskpane/components/chat/MessageBubble.tsx [MOD] Bug #8: null content guard
src/taskpane/components/shared/PromptTemplatesPanel.tsx [MOD] Re-applied nav-bar fix
src/taskpane/components/layout/NavigationTabs.tsx       [MOD] Re-applied nav-bar fix
```
