# 📄 Your Co-Pilot — Comprehensive Page & UI Details

This document provides an exhaustive, technical, and architectural breakdown of every **Page**, **Panel**, **Component**, **Tab**, **Ribbon Command**, and **User Workflow** within **Your Co-Pilot (Office Add-in)**. 

Whether you are a developer extending the codebase, an enterprise admin reviewing security and functionality, or an advanced user wanting to understand every feature, this reference guide covers the entire user interface and interaction lifecycle.

---

## 📋 Table of Contents

1. [Application Overview & Design System](#1-application-overview--design-system)
2. [Office Ribbon & Commands](#2-office-ribbon--commands)
3. [Primary Navigation & Toolbar (`NavigationTabs.tsx`)](#3-primary-navigation--toolbar-navigationtabstsx)
4. [Chat Page & View (`ChatPanel.tsx`)](#4-chat-page--view-chatpaneltsx)
   - [Chat Input Hub (`ChatInput.tsx`)](#chat-input-hub-chatinputtsx)
   - [Message Bubble Rendering (`MessageBubble.tsx`)](#message-bubble-rendering-messagebubbletsx)
5. [Settings Page & Configuration (`SettingsPanel.tsx`)](#5-settings-page--configuration-settingspaneltsx)
   - [API Key Manager (`ApiKeyManager.tsx`)](#api-key-manager-apikeymanagertsx)
   - [Custom Model Manager (`CustomModelManager.tsx`)](#custom-model-manager-custommodelmanagertsx)
   - [Model Selector (`ModelSelector.tsx`)](#model-selector-modelselectortsx)
6. [Automated Copilot Action Engines (`services/office/`)](#6-automated-copilot-action-engines-servicesoffice)
   - [Excel Automation Engine (`ExcelService.ts`)](#excel-automation-engine-excelservicets)
   - [Word Automation Engine (`WordService.ts`)](#word-automation-engine-wordservicets)
   - [PowerPoint Automation Engine (`PowerPointService.ts`)](#powerpoint-automation-engine-powerpointservicets)
7. [Shared UI Components & Utilities (`components/shared/`)](#7-shared-ui-components--utilities-componentsshared)
8. [State Management & Data Flow (`store/AppContext.tsx`)](#8-state-management--data-flow-storeappcontexttsx)

---

## 1. Application Overview & Design System

Your Co-Pilot is built as a responsive, high-performance **Single Page Application (SPA)** embedded directly inside the Microsoft Office Taskpane (Excel, Word, and PowerPoint) on Windows, macOS, and Web browsers.

```
+-----------------------------------------------------------------------+
|                    MICROSOFT OFFICE HOST WINDOW                       |
|                                                                       |
|  +--------------------------------------------+  +-----------------+  |
|  |                                            |  |  TASKPANE       |  |
|  |                                            |  |  Your Co-Pilot  |  |
|  |          ACTIVE WORKBOOK / DOC             |  |                 |  |
|  |                                            |  |  +-----------+  |  |
|  |                                            |  |  | Nav Tabs  |  |  |
|  |                                            |  |  +-----------+  |  |
|  |                                            |  |  |           |  |  |
|  |                                            |  |  | ChatView  |  |  |
|  |                                            |  |  |    or     |  |  |
|  |                                            |  |  | Settings  |  |  |
|  |                                            |  |  |           |  |  |
|  |                                            |  |  +-----------+  |  |
|  |                                            |  |  | ChatInput |  |  |
|  |                                            |  |  +-----------+  |  |
|  +--------------------------------------------+  +-----------------+  |
+-----------------------------------------------------------------------+
```

### Design System Highlights
- **Microsoft Fluent UI v9**: Employs official `@fluentui/react-components` and `@fluentui/react-icons` tokens and components. This ensures a native, visually seamless appearance that blends perfectly into Microsoft 365.
- **Dynamic Theme Mode**: Supports **Light**, **Dark**, and **Auto** modes. When set to Auto, the add-in listens to the host Office application's theme and switches dynamically.
- **Responsive Layout**: Designed to work gracefully across any taskpane width (minimum 300px), automatically adjusting spacing, icon buttons, and text wrapping.

---

## 2. Office Ribbon & Commands

The entry point into Your Co-Pilot begins at the Office Ribbon.

### Ribbon Button Integration (`manifests/*.xml`)
- In Excel, Word, and PowerPoint manifests, a custom Ribbon group titled **Your Co-Pilot** is injected into the **Home Tab**.
- **Button Command**: Clicking the **Your Co-Pilot** button triggers `showTaskpane`, which expands the sidebar UI inside the Office host.
- **Iconography**: Custom high-resolution icons (`16x16`, `32x32`, `80x80`) provide a sharp brand appearance across standard and Retina displays.

### Commands Entry (`src/commands/commands.ts`)
- Registers Office ribbon actions cleanly with `Office.onReady()`.
- Exposes `showTaskpane(event)` globally to handle Office Add-in UI invocations without page reloads.

---

## 3. Primary Navigation & Toolbar (`NavigationTabs.tsx`)

The **Navigation Bar** sits persistently at the top of the taskpane, acting as the main control center for views, context toggling, and conversation management.

```
+-----------------------------------------------------------------------+
|  [💬 Chat]  |  [📄 Context: ON]  [📜 History]  [➕ New]  |  [⚙️ Settings]  |
+-----------------------------------------------------------------------+
```

### Key Functional Sections

| Section | Element | Description |
| :--- | :--- | :--- |
| **Left** | **View Switcher Button** | Switches to the primary Chat interface (`ChatPanel`). Displays subtle highlight styling when active. |
| **Middle** | **Document Context Toggle** | A live toggle button (`DocumentTextRegular`) that controls whether current cell selection or document text is appended to the AI prompt. Colored **Green (`#76B900`)** when **ON**. |
| **Middle** | **Chat History Menu** | A dropdown menu (`MenuPopover`) listing all saved chat threads. Shows titles, timestamps, and active conversation highlight. |
| **Middle** | **New Chat Button** | Instantly initializes a fresh conversation thread while preserving previous history. |
| **Middle** | **Delete Chat Button** | Deletes the currently active conversation with a confirm toast notification. |
| **Right** | **Settings Button** | Switches the taskpane view to the **Settings Panel** (`SettingsPanel.tsx`). |

---

## 4. Chat Page & View (`ChatPanel.tsx`)

The **Chat Page** is the heart of Your Co-Pilot. It provides a real-time, streaming conversational interface equipped with Office automation awareness.

### Layout Structure
- **Header Banner**: Displays the current active model, provider badge, and context inclusion status.
- **Message List Area**: A scrollable viewport rendering conversation bubbles, empty states, and loading animations.
- **Bottom Anchor**: Holds the responsive `ChatInput` component fixed at the bottom of the taskpane.

---

### Chat Input Hub (`ChatInput.tsx`)

```
+-----------------------------------------------------------------------+
|  Selected Range: A1:D15 (Sheet1)                   [💡 Templates (50+)] |
+-----------------------------------------------------------------------+
|  Ask AI to analyze, write formulas, or format...                      |
|                                                                       |
|  [                                                 ]       [ 🚀 Send ]|
+-----------------------------------------------------------------------+
```

#### Features & Mechanics
- **Auto-expanding Textarea**: Expands dynamically as the user types multi-line prompts.
- **Keyboard Shortcuts**:
  - `<Enter>`: Submits the message immediately.
  - `<Shift> + <Enter>`: Inserts a line break.
- **Live Context Banner**: Automatically queries `ExcelService`, `WordService`, or `PowerPointService` to show exactly what data range or text is currently selected (e.g., `"Selected: A1:G40 on Financials"`).
- **Prompt Templates Drawer (`💡 Templates`)**:
  - Provides **50+ curated prompt templates** categorized by host app (Excel, Word, PowerPoint).
  - Categorized into *Formulas*, *Data Cleaning*, *Financial Analysis*, *Summarization*, *VBA/Macros*, and *Visualizations*.
  - Clicking a template auto-populates the input area or submits directly.
- **Streaming Controller**: During active AI text generation, the Submit button transforms into a red **Stop / Abort** button, allowing users to cancel streaming requests instantly via `AbortController`.

---

### Message Bubble Rendering (`MessageBubble.tsx`)

Each message in the conversation is rendered inside a rich, structured `MessageBubble` that supports Markdown, syntax highlighting, and visual indicators.

#### Core Visual & Interactive Components
1. **Role Identification**:
   - **User Bubbles**: Right-aligned with a subtle brand tint background.
   - **Assistant Bubbles**: Left-aligned with a clean card background, avatar badge, and AI model metadata.
2. **Real-time Typewriter Streaming**:
   - Token-by-token rendering with a smooth animated cursor while the stream is active.
3. **Rich Markdown & Syntax Highlighting**:
   - Powered by `MarkdownRenderer.tsx` and `CodeBlock.tsx`.
   - Renders GitHub-flavored markdown (GFM): tables, lists, bold text, blockquotes, and code blocks.
   - Code snippets (VBA, Excel Formulas, Python, JavaScript, SQL) feature line numbers, syntax coloring via Prism, and a **1-Click Copy to Clipboard** button.
4. **Context Pill Indicator**:
   - Displays a badge (`"📄 Included A1:C20"`) on user messages when spreadsheet/document context was sent to the AI.
5. **Office Action Badges**:
   - When the AI emits invisible automation commands (`<EXCEL_CMD>`, `<WORD_CMD>`, `<PPT_CMD>`), the message bubble parses them and displays an elegant action execution summary badge (e.g., `⚡ Created Chart 'Sales Summary'`).

---

## 5. Settings Page & Configuration (`SettingsPanel.tsx`)

The **Settings Page** is a centralized control panel for managing multi-provider API keys, custom models, UI themes, and privacy rules.

```
+-----------------------------------------------------------------------+
|  ⚙️ Settings & AI Providers                          [Back to Chat ✖] |
+-----------------------------------------------------------------------+
|  🤖 AI Configuration                                                  |
|  +-----------------------------------------------------------------+  |
|  | Provider: [ NVIDIA NIM (Free Credits)            ▼ ]            |  |
|  | Model:    [ Llama 3.1 70B Instruct              ▼ ]            |  |
|  +-----------------------------------------------------------------+  |
|                                                                       |
|  🔑 API Keys & Authentication   [ Expand / Collapse ]                  |
|  🛠️ Custom Models Manager      [ Add Custom Model  ]                  |
|  🎨 UI Theme & Display          [ Light / Dark / Auto ]               |
|  🔒 Privacy & Context Controls  [ Max Cells / Characters ]            |
+-----------------------------------------------------------------------+
```

### API Key Manager (`ApiKeyManager.tsx`)
- **Seven Native Providers Supported**:
  1. **NVIDIA NIM** (`build.nvidia.com` - Free credits, Llama 3.1, Mixtral, Nemotron)
  2. **Google Gemini** (`aistudio.google.com` - Generous 1M context free tier)
  3. **Groq** (`console.groq.com` - Ultra-fast Llama 3.1 inference)
  4. **OpenRouter** (`openrouter.ai` - Free & paid multi-model access)
  5. **OpenAI** (`platform.openai.com` - GPT-4o, GPT-4o-mini)
  6. **Anthropic** (`console.anthropic.com` - Claude 3.5 Sonnet, Haiku)
  7. **Ollama** (`ollama.com` - 100% free, offline, private local models)
- **Security & Privacy**: Keys are stored strictly in the browser's `localStorage` per provider. They are **never** transmitted to any intermediary server.
- **Live Key Testing**: Features a **Test Connection** button that sends a lightweight test ping to verify key validity before saving.

---

### Custom Model Manager (`CustomModelManager.tsx`)
- Allows users to add arbitrary AI models from any OpenAI-compatible endpoint or OpenRouter catalog.
- **Configurable Attributes**:
  - `Display Name`: Friendly name shown in dropdowns.
  - `Model ID`: Exact backend identifier (e.g., `meta-llama/llama-3.1-405b-instruct`).
  - `Provider`: Associated provider routing.
  - `Custom API Key` & `Custom Base URL`: Override default endpoints for private corporate gateways or self-hosted LLMs.

---

### Model Selector (`ModelSelector.tsx`)
- An intelligent dropdown selector that groups models by provider.
- Displays metadata badges for each model, including **Context Window** size (e.g., `128k`, `1M`) and **Max Tokens** limits.

---

### UI & Privacy Settings
- **Theme Mode**: Switch between `Light`, `Dark`, and `Auto` (Office system sync).
- **Font Size Adjuster**: Toggle between `Small`, `Medium`, and `Large` UI typography.
- **Context Limits**:
  - `includeContextByDefault`: Set whether selection is attached automatically.
  - `maxContextCells`: Guardrail limiting how many Excel cells are serialized (default: 500 cells) to prevent token overflow.
  - `maxContextCharacters`: Safeguard against sending overly large text selections.

---

## 6. Automated Copilot Action Engines (`services/office/`)

Your Co-Pilot is distinguished by its ability to execute **autonomous document manipulations**. The AI emits structured JSON commands inside `<EXCEL_CMD>`, `<WORD_CMD>`, or `<PPT_CMD>` tags, which are intercepted by specialized service engines.

---

### Excel Automation Engine (`ExcelService.ts`)

Implements **40+ direct Excel.js API operations**, plus an unconstrained JavaScript evaluation sandbox:

| Category | Supported Actions | Description |
| :--- | :--- | :--- |
| **Writing & Formulas** | `write_cell`, `write_formula`, `write_range` | Writes single cells, formulas (including dynamic arrays like `FILTER`, `XLOOKUP`), and 2D arrays. |
| **Visualization** | `create_chart`, `delete_chart`, `format_chart`, `add_sparklines` | Generates column, line, bar, pie, and scatter charts; adds inline cell sparklines. |
| **Data Analysis** | `create_pivot_table`, `create_table`, `sort_range`, `apply_filter`, `clear_filter` | Inserts structured Excel tables and PivotTables; applies multi-column sorting and filtering. |
| **Formatting & Styles** | `format_range`, `add_conditional_formatting`, `highlight_duplicates`, `highlight_top_bottom` | Modifies fonts, colors, borders, and number formats; applies color scales and conditional highlights. |
| **Data Cleaning** | `remove_duplicates`, `trim_whitespace`, `change_case`, `remove_blank_rows` | Purges duplicate rows, trims excess spaces, standardizes text casing, and cleans blanks. |
| **Worksheet Layout** | `add_sheet`, `delete_sheet`, `insert_range`, `delete_range`, `merge_cells`, `group_data`, `ungroup_data` | Manages sheet tabs, inserts/deletes row/column blocks, merges headers, and outlines row groups. |
| **Navigation & Panes** | `freeze_panes`, `unfreeze_panes`, `autofit_columns`, `autofit_rows` | Locks header rows/columns and auto-fits cell dimensions to text content. |
| **Named Ranges & CSV** | `create_named_range`, `delete_named_range`, `search_data`, `get_all_objects`, `get_range_csv` | Creates named references, performs regex/literal searches, and reads ranges as token-efficient CSV. |
| **Escape Hatch** | `eval_js` | **SES Sandboxed Execution**: Executes dynamic, arbitrary `Excel.run(async (context) => { ... })` scripts on the fly. |

---

### Word Automation Engine (`WordService.ts`)

Provides automated drafting and editing capabilities using `Word.run`:

- **`insert_paragraph`**: Inserts formatted text before, after, at the start, or at the end of the selection.
- **`insert_table`**: Builds multi-row Word tables with styled headers automatically.
- **`format_text`**: Modifies bold, italics, font sizing, and font colors across selected text.
- **`apply_style`**: Applies built-in Word styles (`Heading1`, `Heading2`, `Title`, `Normal`).
- **`clear_formatting`**: Strips messy copy-pasted styles back to clean body text.
- **`search_replace` & `highlight_search`**: Mass find-and-replace or color-highlights target search phrases.
- **`get_structure`**: Reads document outline, heading hierarchy, and word/paragraph statistics.
- **`eval_js`**: Sandboxed execution of arbitrary Word.js API commands.

---

### PowerPoint Automation Engine (`PowerPointService.ts`)

Enables presentation construction using `PowerPoint.run`:

- **`add_slide`**: Inserts blank slides or new slides with standard layout structures.
- **`add_textbox`**: Places positioned text boxes with custom typography onto slides.
- **`add_shape` & `format_shape`**: Inserts geometric shapes (`Rectangle`, `Oval`, `Line`) and configures fill/line colors.
- **`set_slide_notes`**: Injects comprehensive speaker notes into the presenter notes pane.
- **`get_shapes` & `delete_slide`**: Inspects all shapes on a slide or deletes slides by index.
- **`eval_js`**: Sandboxed execution of arbitrary PowerPoint.js API commands.

---

## 7. Shared UI Components & Utilities (`components/shared/`)

- **`CodeBlock.tsx`**: Renders code snippets with syntax highlighting via `prism-react-renderer`. Includes a header bar showing the programming language and a **Copy** button.
- **`CopyButton.tsx`**: A reusable micro-button that copies text to the clipboard and displays a temporary checkmark (`✓ Copied!`).
- **`MarkdownRenderer.tsx`**: A robust wrapper around `react-markdown` and `remark-gfm` configured with custom renderers to match Fluent UI typography and table borders.
- **`NotificationToast.tsx`**: Floating alert bar that renders transient feedback messages (`Success`, `Error`, `Warning`, `Info`) with auto-dismiss timers.
- **`EmptyState.tsx` & `LoadingDots.tsx`**: Provide intuitive, branded empty states and animated typing indicators when waiting for AI responses.

---

## 8. State Management & Data Flow (`store/AppContext.tsx`)

State across the add-in is managed via React Context (`AppContext.tsx`) combined with custom hooks:

```
                  +-----------------------------------+
                  |        AppContext Provider        |
                  |  (Active View, Host, Settings)    |
                  +-----------------+-----------------+
                                    |
          +-------------------------+-------------------------+
          |                                                   |
          v                                                   v
+-------------------+                               +-------------------+
|    useChat.ts     |                               |  useSettings.ts   |
| - Conversations   |                               | - API Keys        |
| - Streaming State |                               | - Custom Models   |
| - AI Invocation   |                               | - Theme & Limits  |
+---------+---------+                               +---------+---------+
          |                                                   |
          +-------------------------+-------------------------+
                                    |
                                    v
                  +-----------------------------------+
                  |        ProviderFactory.ts         |
                  |  (NVIDIA, Gemini, Groq, OpenAI,   |
                  |   Anthropic, OpenRouter, Ollama)  |
                  +-----------------------------------+
```

### Core Hooks
- **`useChat(hostApp)`**: Orchestrates conversation threads, builds prompt templates (`buildSystemPrompt`), invokes `ProviderFactory`, and parses `<EXCEL_CMD>`, `<WORD_CMD>`, and `<PPT_CMD>` blocks from AI outputs.
- **`useSettings()`**: Manages persistence to `localStorage` and provides settings getters/setters.
- **`useTheme()`**: Resolves current host theme colors and applies Fluent UI tokens.
- **`useAI()`**: Connects view components directly to provider test pings and model availability checks.

---

## 9. Verification & Architectural Safety

- **No Remote Telemetry**: Your Co-Pilot executes entirely locally inside the Office taskpane WebView.
- **Sandboxed Execution**: Any dynamic scripts run via `eval_js` execute within Microsoft Office's native `RequestContext` (`Excel.run`, `Word.run`, `PowerPoint.run`), preventing access to external system resources.
- **Cross-Platform Compatibility**: Fully tested and compatible with Excel/Word/PowerPoint 2016, 2019, 2021, Microsoft 365 Desktop, and Office Online.

---

<p align="center">
  <strong>Built with ❤️ for the Open-Source Microsoft Office & AI Communities</strong>
</p>
