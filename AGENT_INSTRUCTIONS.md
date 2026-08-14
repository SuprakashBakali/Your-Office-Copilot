# Agent Instructions — Whole-Document Awareness + Attachments + Status Fix

## What was fixed (4 major changes)

### 1. Response-finished indicator (ChatInput.tsx)
**Problem:** User couldn't tell when the AI finished responding.
**Fix:** Added a `justFinished` state that shows a green "✓ Done" indicator for 3 seconds after streaming ends. The indicator transitions: Ready → Generating... (spinner) → ✓ Done (3s) → Ready.

### 2. Whole-document context (ExcelService, WordService, PowerPointService)
**Problem:** Context was selection-only or active-sheet-only. The AI couldn't see the whole workbook/document/presentation.
**Fix:**
- **Excel:** `getContextForAI()` already lists all sheets (from previous fix). Now also includes named ranges.
- **Word:** `getContextForAI()` now reads FULL document text + all headings + paragraph/table counts + selected text. Increased maxChars to 8000.
- **PowerPoint:** `getContextForAI()` now reads ALL slides via PowerPointApi 1.4+ — iterates every slide, every shape, extracts all text. Falls back to old method on older API.

### 3. PowerPoint data insertion fixed (PowerPointService.ts)
**Problem:** `addTextbox` didn't set position/size — text was invisible or off-slide.
**Fix:** `addTextbox` now sets `left=50, top=100, width=600, height=100` so text is always visible. Also added optional `slideIndex` parameter.

### 4. Attachment button (ChatInput.tsx + useChat.ts + useAI.ts)
**Problem:** No way to attach images/PDFs for the AI to parse.
**Fix:** Added a paperclip button that accepts images, PDFs, Excel/Word/PPT files. Images are sent as multimodal content (`image_url` with base64 data URL) so vision-capable models can see them. Non-image files are noted as text references.

## Files changed (7)

1. **ChatInput.tsx** — Attachment button + file picker + `justFinished` status + attachment chips
2. **ChatPanel.tsx** — Pass attachments through to sendChatMessage
3. **useChat.ts** — Accept attachments param, build multimodal content for AI
4. **useAI.ts** — buildRequestOptions handles `_multimodalContent` (image_url array)
5. **WordService.ts** — getContextForAI reads full document structure + text
6. **PowerPointService.ts** — getContextForAI reads ALL slides; addTextbox sets position/size
7. **prompts.ts** — Word + PPT prompts updated with DOCUMENT/PRESENTATION AWARENESS + workflow examples

## Verification

```
tsc --noEmit              → 0 errors
eslint --max-warnings 0   → 0 errors, 0 warnings
webpack --mode production  → succeeds, 962 KB
```

## How to deploy

```bash
unzip whole-doc-attachments.zip -d YOUR_REPO
cd YOUR_REPO
npm install
git add -A
git commit -m "feat: whole-document context + attachments + response-finished indicator

- ChatInput: '✓ Done' status shows for 3s after streaming ends
- ChatInput: paperclip button for image/PDF/file attachments
- Attachments: images sent as multimodal content (image_url) to AI
- WordService: getContextForAI reads full doc text + headings + structure
- PowerPointService: getContextForAI reads ALL slides via PPT API 1.4
- PowerPointService: addTextbox sets position/size (was invisible)
- prompts.ts: Word + PPT prompts have DOCUMENT/PRESENTATION AWARENESS
- prompts.ts: batch workflow examples (emit many CMD blocks at once)"
git push origin main
```
