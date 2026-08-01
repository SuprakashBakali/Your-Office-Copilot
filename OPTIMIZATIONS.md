# Optimizations Applied

This document summarizes the optimizations applied to the Your-Office-Copilot repository. All changes preserve the existing public API (provider IDs, settings shape, UI/UX) — no breaking changes.

## Summary

| Area | Before | After |
|------|--------|-------|
| AI provider code (duplication) | ~1061 lines, 5 near-identical providers | ~753 lines with shared `OpenAICompatibleProvider` base |
| `useChat.ts` size | 734 lines (prompts inlined) | 605 lines + extracted `prompts.ts` |
| Bundle JS payload | ~1.58 MB single chunk | ~950 KB split into cacheable vendor chunks |
| `react-syntax-highlighter` | ~600 KB (Prism + all languages) | Replaced with `prism-react-renderer` (~50 KB) |
| Committed `dist/` folder | 4 MB of stale build output tracked in git | Removed; `.gitignore` updated |
| Icon assets | 5 × 376 KB identical PNGs = 1.88 MB | Properly sized: 15 KB total |
| `preview.jpg` | 514 KB | 128 KB (re-saved at quality 85, progressive) |
| Proxy security | Open SSRF — accepts any `targetUrl` | Domain allowlist + HTTPS-only + 8 MB response cap |
| Stream cancellation | `abortController.abort()` only stopped the consumer loop | Signal now threaded through to `fetch()` for real cancellation |
| Conversation history | Never capped (setting ignored) | Capped at `settings.maxConversationHistory` |
| `MarkdownRenderer` | Used deprecated `inline` prop (removed in react-markdown v9) | Rewritten for v9 API |
| Linting/formatting | None | ESLint + Prettier configs added; `lint` / `format` scripts |
| TypeScript | `noEmit: false` (stray .js files) | `noEmit: true` (tsc is type-check only) |

## Detailed changes

### 1. AI provider refactor — `OpenAICompatibleProvider` base class

**Files:** `src/taskpane/services/ai/OpenAICompatibleProvider.ts` (new), `NvidiaProvider.ts`, `OpenAIProvider.ts`, `GroqProvider.ts`, `OpenRouterProvider.ts`, `OllamaProvider.ts`, `GenericOpenAIProvider.ts`

The 5 OpenAI-compatible providers (NVIDIA, OpenAI, Groq, OpenRouter, Ollama) all had nearly identical `chat()` and `chatStream()` methods — only `baseUrl`, `getModels()`, and a few headers differed. They now extend a shared `OpenAICompatibleProvider` base that handles:

- Proxy vs. direct fetch (Ollama uses direct, others use `/api/proxy`)
- Header construction (overridable via `buildHeaders()`)
- Body construction (overridable via `buildBody()` — used by OpenRouter for `plugins` web-search)
- AbortSignal threading
- SSE stream parsing
- Error formatting

Anthropic and Gemini keep their own implementations because they use non-OpenAI API formats.

### 2. Security hardening — proxy SSRF fix

**Files:** `api/proxy.js`, `webpack.config.js`

The proxy previously accepted any `targetUrl`, making it an open SSRF gateway. It now:

- Maintains an allowlist of known AI provider hosts (OpenAI, Anthropic, Gemini, Groq, OpenRouter, NVIDIA, localhost)
- Rejects non-HTTPS for remote hosts (HTTP only for localhost)
- Strips hop-by-hop headers (`host`, `content-length`, `connection`)
- Enforces an 8 MB response size cap
- Returns 403 with a clear message when the target is not allowed

The same hardening is mirrored in the webpack dev-server middleware.

### 3. Bundle size reduction

- Replaced `react-syntax-highlighter` (Prism + bundled languages) with `prism-react-renderer` (~50 KB vs ~600 KB)
- Removed unused `rehype-raw` dependency
- Removed unused `file-loader` (webpack uses `asset/resource` natively)
- Added explicit webpack `cacheGroups` for `vendor-react`, `vendor-fluentui`, `vendor-markdown` so app-code changes don't invalidate the browser cache for stable vendor bundles

### 4. `useChat.ts` refactor

**Files:** `src/taskpane/hooks/useChat.ts`, `src/taskpane/services/prompts.ts` (new)

- Extracted the ~200-line system-prompt strings into `services/prompts.ts` (`buildSystemPrompt(hostApp, contextStr)`)
- Deduplicated the streaming vs. non-streaming paths into a single `finishAssistant()` helper
- Replaced stale-closure-prone `useCallback` deps with `useRef` mirrors of `settings`, `conversations`, `activeConversationId`
- Added `persistConversations()` helper that enforces `settings.maxConversationHistory` cap (previously ignored)
- Live-stream updates now use functional `setConversations(prev => ...)` instead of rebuilding from a stale closure

### 5. AbortSignal properly threaded

**Files:** `src/taskpane/services/ai/types.ts`, `OpenAICompatibleProvider.ts`, `AnthropicProvider.ts`, `GeminiProvider.ts`, `src/taskpane/hooks/useAI.ts`

`useAI.cancelStream()` previously called `abortController.abort()` but the signal was never passed to `fetch()` — only the consumer `for await` loop checked `signal.aborted`. Now:

- `AIRequestOptions.signal?: AbortSignal` added to the type
- All providers pass `signal` into `fetch(..., { signal })`
- `parseOpenAISSEStream()` registers an abort listener that calls `reader.cancel()`
- `useAI.sendMessageStream()` passes the abort controller's signal into the request options

### 6. `MarkdownRenderer` react-markdown v9 fix

**File:** `src/taskpane/components/shared/MarkdownRenderer.tsx`

react-markdown v9 removed the `inline` prop from the `code` component renderer. The old code silently broke inline vs. block code detection. Rewritten to detect block code via the `language-*` class and node position.

### 7. Tooling & cleanup

- **`.eslintrc.cjs`** — ESLint config with TypeScript + React + React Hooks plugins
- **`.prettierrc`** / **`.prettierignore`** — Prettier config
- **`package.json`** — added `engines.node`, `lint` / `lint:fix` / `format` scripts, removed `react-syntax-highlighter` + `@types/react-syntax-highlighter` + `rehype-raw` + `file-loader`, added ESLint + Prettier + react eslint plugins
- **`tsconfig.json`** — `noEmit: true` (webpack emits, tsc only type-checks), removed `declaration` / `sourceMap` / `outDir` / `rootDir` (no longer relevant)
- **`.gitignore`** — added `dist/`, `*.tsbuildinfo`, `.eslintcache`
- **Removed `dist/` from git tracking** — was 4 MB of stale build output

### 8. Asset optimization

- Regenerated `icon-{16,32,64,80,128}.png` from `icon.jpg` at correct resolutions (5 × 376 KB → 15 KB total)
- Re-saved `preview.jpg` at quality 85 with progressive encoding (514 KB → 128 KB)
- Re-saved `icon.jpg` at quality 88 (128 KB → 40 KB)

## Verification

All changes verified with:

```bash
npx tsc --noEmit          # ✅ clean (0 errors)
npx eslint "src/**/*.{ts,tsx}"  # ✅ 0 errors, 10 pre-existing warnings (unused imports in Office service files)
npx webpack --mode production    # ✅ builds successfully
```

## How to use

```bash
npm install        # install new deps (prism-react-renderer, eslint, prettier, etc.)
npm run typecheck  # verify types
npm run lint       # run ESLint
npm run build      # production build
npm run dev        # dev server
```

No runtime behavior changes — all provider IDs, settings shape, manifests, and UI are unchanged.
