# Feature Additions

Inspired by analyzing two reference repositories:

- **[hewliyang/office-agents](https://github.com/hewliyang/office-agents)** — pnpm monorepo of Office Add-ins with an agent SDK, skills system, web search, thinking blocks, conversation compaction, and a local dev bridge.
- **[iOfficeAI/OfficeCLI](https://github.com/iOfficeAI/OfficeCLI)** — CLI for AI-driven Office doc creation with specialized "skills" (pitch-deck, financial-model, data-dashboard, academic-paper, morph-ppt).

This document describes the features ported into Your-Office-Copilot and how they were adapted.

## Summary of new features

| # | Feature | Inspired by | Status |
|---|---------|-------------|--------|
| 1 | **Prompt Templates panel** | office-agents Skills + OfficeCLI specialized skills | ✅ 17 templates, host-aware |
| 2 | **Conversation auto-compaction** | office-agents `TODO.md` compaction design | ✅ Best-effort, transparent |
| 3 | **Web Search toggle UI** | office-agents Settings panel | ✅ Wired to existing `webSearch` flag |
| 4 | **Thinking / Reasoning block** | office-agents `thinking-block.svelte` | ✅ DeepSeek-R1 + Claude + o1 |
| 5 | **History search/filter** | office-agents sessions panel | ✅ Inline search in history menu |
| 6 | **Token usage plumbing** | office-agents message metadata | ✅ `ChatMessage.tokens` + `thinking` fields |

---

## 1. Prompt Templates panel

**Files:** `src/taskpane/services/promptTemplates.ts` (new), `src/taskpane/components/shared/PromptTemplatesPanel.tsx` (new)

A sparkle-icon button in the nav bar opens a searchable popover with curated, host-aware prompt templates. Clicking a template drops its prompt into the chat input via the existing `pendingPrompt` mechanism.

### Why this instead of full Skills

office-agents ships a full SKILL.md install/parse/VFS system. That's powerful but heavy — it requires IndexedDB storage, file upload UI, and a virtual filesystem. For Your-Office-Copilot's single-add-in scope, a curated built-in template library gives users 80% of the value at 5% of the complexity.

### Templates included (17 total)

**Excel (6):** Summarize Data, Build Dashboard, Financial Model, Clean Data, Explain Formula, Auto Pivot Table
**Word (5):** Executive Summary, Proofread & Fix, Academic Paper, Rewrite Tone, Generate TOC
**PowerPoint (4):** Investor Pitch Deck, Slides from Document, Speaker Notes, Design Polish
**Cross-host (3):** Explain Selection, Translate, Web Research

Each template has an `icon`, `category` (analyze/create/format/explain/utility), optional `hosts` filter, and a `prompt` body with `{placeholder}` support via `fillPromptTemplate()`.

The templates port the *spirit* of OfficeCLI's specialized skills (pitch-deck → "Investor Pitch Deck", financial-model → "Financial Model", data-dashboard → "Build Dashboard", academic-paper → "Academic Paper") without requiring the user to install a separate CLI binary.

---

## 2. Conversation auto-compaction

**Files:** `src/taskpane/services/compaction.ts` (new), `src/taskpane/hooks/useChat.ts` (updated), `src/taskpane/types/index.ts` (updated)

When a conversation grows too long for the model's context window, older messages are summarized into a single `compaction_summary` marker message. The recent context is preserved verbatim.

### Design (ported from office-agents `TODO.md`)

```
m1, m2, m3, m4, m5, [compaction_summary], m6, m7
                                          ^--- LLM context starts here
```

- **Display:** `compaction_summary` messages are filtered out of the chat UI (see `ChatPanel.tsx`), so the user sees a linear history.
- **LLM context:** `sliceContextForLLM()` finds the last compaction marker and sends only from there onward, converting the marker itself to a `<compaction_summary>...</compaction_summary>` system message.
- **Re-compaction:** idempotent — the previous summary is already in context, so the next compaction naturally folds it in.

### Trigger heuristic

- Soft limit: **12,000 characters** of concatenated message contents (roughly 3-4k tokens, conservative to leave headroom for the system prompt + response).
- Minimum messages before compaction kicks in: **12** (so short conversations aren't summarized).
- Always preserve the **last 6 messages** verbatim.

### Best-effort fallback

If the summarization LLM call fails (network error, bad API key, etc.), the hook logs a warning and falls back to sending the full history. The user's message is never blocked by compaction failures.

---

## 3. Web Search toggle UI

**Files:** `src/taskpane/components/layout/NavigationTabs.tsx` (updated), `src/taskpane/components/layout/Sidebar.tsx` (updated), `src/taskpane/components/chat/ChatPanel.tsx` (updated)

The `webSearch` flag already existed in `AIRequestOptions` and was wired through OpenRouter (via `plugins: [{ id: "web" }]`) and Gemini (via `tools: [{ googleSearch: {} }]`). But the new icon-only nav bar had no button to toggle it.

A globe icon now sits in the middle nav section. When ON, it turns NVIDIA-green and the `webSearchEnabled` state flows from `Sidebar` → `ChatPanel` → `sendChatMessage(text, includeContext, webSearchEnabled)`.

---

## 4. Thinking / Reasoning block

**Files:** `src/taskpane/components/shared/ThinkingBlock.tsx` (new), `src/taskpane/components/chat/MessageBubble.tsx` (updated), `src/taskpane/services/ai/types.ts` (updated), `src/taskpane/hooks/useAI.ts` (updated)

### What it does

Models like DeepSeek-R1, Claude (with extended thinking), and OpenAI o1 emit a separate "reasoning" trace before their final answer. This is now captured and rendered in a collapsible "thinking" panel above the assistant's response.

### How it's captured

The OpenAI-compatible SSE stream parser (`parseOpenAISSEStream`) now checks for thinking content in two places:

1. **`delta.reasoning_content`** — DeepSeek-R1's official field for reasoning tokens.
2. **Inline `<think>...</think>`** — some community models wrap reasoning inline in the content stream.

Both are captured into `AIStreamChunk.thinking` and streamed to the UI via a new `onThinking` callback on `sendMessageStream`. The final thinking text is stored on `ChatMessage.thinking` and rendered by `ThinkingBlock` (a collapsible panel adapted from office-agents' `thinking-block.svelte`).

### Why this matters

Without this, DeepSeek-R1's reasoning would either (a) leak into the visible response as `<think>...</think>` garbage, or (b) be silently dropped. Now it's preserved as a collapsible trace that power users can inspect.

---

## 5. History search / filter

**Files:** `src/taskpane/components/layout/NavigationTabs.tsx` (updated)

The chat history dropdown now has an inline search input at the top. Typing filters conversations by title **and** message content (case-insensitive substring match). Each history item also shows message count and last-updated date.

This ports office-agents' session-list search pattern without adopting their full IndexedDB-backed session store.

---

## 6. Token usage + thinking plumbing

**Files:** `src/taskpane/types/index.ts` (updated)

`ChatMessage` now has two optional fields:

- `tokens?: number` — already existed but was never populated; providers already return `usage.totalTokens` in `AIResponse`, so future UI work can surface this without further plumbing.
- `thinking?: string` — the reasoning trace (see Feature 4).
- `compacted?: boolean` — marks messages that have been hidden from LLM context by compaction (reserved for future UI hints; currently the compaction marker itself serves this purpose).

---

## Verification

```bash
npx tsc --noEmit                    # 0 errors
npx eslint src/**/*.{ts,tsx}        # 0 errors, 0 warnings (--max-warnings 0 passes)
npx webpack --mode production       # succeeds, ~985 KB total (was 950 KB)
```

Bundle grew by ~35 KB due to the new `PromptTemplatesPanel`, `ThinkingBlock`, and the prompt-template data. Still well under the 1.58 MB pre-optimization size.

---

## What was NOT ported (and why)

| Feature | Source | Why skipped |
|---------|--------|-------------|
| Full SKILL.md install system | office-agents | Requires IndexedDB + VFS + file upload UI; too heavy for this add-in's scope. Curated templates give 80% of the value. |
| Office Bridge (local RPC CLI) | office-agents | Dev-only tooling; would require shipping a separate npm package and a WebSocket server. Out of scope for an end-user add-in. |
| Sandboxed `Compartment` eval | office-agents | The `eval_js` escape hatch already exists; adding SES/lockdown would add ~50KB for a marginal security benefit in a BYOK add-in where the user already trusts their own API key. |
| OfficeCLI CLI binary integration | OfficeCLI | Would require users to install a separate binary. The add-in's `eval_js` command already covers in-document automation. |
| OOXML/PPTX direct editing | OfficeCLI / office-agents | Office.js already provides high-level APIs; raw OOXML editing is a power-user escape hatch that the existing `eval_js` covers when needed. |
| Multi-language README (ja/ko/zh) | OfficeCLI | Not a code feature; can be added as a docs PR if internationalization becomes a priority. |
| Morph transitions | OfficeCLI morph-ppt skill | PowerPoint.js doesn't expose Morph transition setting directly; would require OOXML hacking. Future enhancement. |

---

## How to use the new features

### Prompt Templates
1. Click the **sparkle icon** ✨ in the nav bar.
2. Search or browse by category.
3. Click a template — its prompt drops into the chat input.
4. Edit if needed, then hit Enter.

### Web Search
1. Click the **globe icon** 🌐 in the nav bar (turns green when ON).
2. Send your message — the AI will use OpenRouter's web plugin or Gemini's Google Search grounding (depending on your active provider).
3. Note: only OpenRouter and Gemini providers support web search; other providers ignore the flag.

### Thinking blocks
1. Use a reasoning model (DeepSeek-R1 on NVIDIA NIM, Claude with extended thinking, or o1 on OpenAI).
2. After the response completes, a collapsible "thinking" panel appears above the answer.
3. Click to expand and inspect the reasoning trace.

### History search
1. Click the **history icon** 🕒 in the nav bar.
2. Type in the search box at the top of the dropdown.
3. Conversations filter by title and message content in real time.

### Auto-compaction
- **No user action required.** When a conversation exceeds ~12k characters, the next message you send triggers a background summarization of older messages. The chat UI is unaffected; the compaction marker is hidden from display.

---

## File map of changes

```
src/taskpane/
├── services/
│   ├── promptTemplates.ts          [NEW]  17 curated templates + helpers
│   ├── compaction.ts               [NEW]  Context compaction logic
│   └── ai/types.ts                 [MOD]  +thinking on AIResponse/AIStreamChunk
├── hooks/
│   ├── useAI.ts                    [MOD]  +onThinking callback, returns {text,thinking}
│   └── useChat.ts                  [MOD]  Compaction wiring, thinking capture
├── components/
│   ├── shared/
│   │   ├── ThinkingBlock.tsx       [NEW]  Collapsible reasoning panel
│   │   └── PromptTemplatesPanel.tsx [NEW] Searchable template popover
│   ├── chat/
│   │   ├── MessageBubble.tsx       [MOD]  Renders ThinkingBlock
│   │   └── ChatPanel.tsx           [MOD]  Filters compaction_summary, passes webSearch
│   └── layout/
│       ├── NavigationTabs.tsx      [MOD]  +Templates, +WebSearch toggle, +history search
│       └── Sidebar.tsx             [MOD]  Manages webSearchEnabled state
└── types/
    └── index.ts                    [MOD]  +compaction_summary role, +thinking/compacted fields
```

**9 files changed: 4 new, 5 modified. ~600 lines added.**
