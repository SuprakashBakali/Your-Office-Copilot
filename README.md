# Your Co-Pilot — Open-Source Microsoft Excel AI Add-in

**The Ultimate Autonomous AI Copilot for Microsoft Excel — Powered by 7 AI Providers with Real-Time Spreadsheet Manipulation**

[![Quick Start](https://img.shields.io/badge/Quick_Start-↓-76B900?style=for-the-badge&logo=rocket)](#-step-by-step-installation--quick-start)
[![Architecture](https://img.shields.io/badge/Architecture-Documentation-blueviolet?style=for-the-badge&logo=read-the-docs)](#-deep-dive-documentation)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue?style=for-the-badge)](LICENSE)
![React 18](https://img.shields.io/badge/React-18.3-61DAFB?style=for-the-badge&logo=react&logoColor=white)
![TypeScript 5.5](https://img.shields.io/badge/TypeScript-5.5-3178C6?style=for-the-badge&logo=typescript&logoColor=white)
![Fluent UI v9](https://img.shields.io/badge/Fluent_UI-v9-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)
![Excel.js Add-in](https://img.shields.io/badge/Excel.js-Add--in-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)

---

## 📋 Table of Contents

1. [What is Your Co-Pilot for Excel?](#-what-is-your-co-pilot-for-excel)
2. [Why Choose Your Co-Pilot Over Paid Alternatives?](#-why-choose-your-co-pilot-over-paid-alternatives)
3. [40+ Excel Autonomous Actions & Sandboxing](#-40-excel-autonomous-actions--sandboxing)
4. [7 Supported AI Providers (Free & Paid)](#-7-supported-ai-providers-free--paid)
5. [Intelligent Chat & Context Experience](#-intelligent-chat--context-experience)
6. [Step-by-Step Installation & Quick Start](#-step-by-step-installation--quick-start)
7. [Technical Architecture & Project Structure](#️-technical-architecture--project-structure)
8. [Development & Validation Commands](#️-development--validation-commands)
9. [Deep-Dive Documentation](#-deep-dive-documentation)
   - [Application Overview & Design System](#1-application-overview--fluent-ui-v9-design-system)
   - [Excel Ribbon Integration](#2-excel-ribbon-integration)
   - [Primary Navigation & Conversation Manager](#3-primary-navigation--conversation-manager)
   - [Chat Panel & Live Range Context](#4-chat-panel--live-range-context)
   - [Settings & Multi-Provider API Configuration](#5-settings--multi-provider-api-configuration)
   - [Excel Automation Engine (ExcelService.ts)](#6-excel-automation-engine-excelservicets)
   - [Shared UI Components & Utilities](#7-shared-ui-components--utilities)
   - [State Management & Data Flow](#8-state-management--data-flow)
10. [Security, Privacy & Enterprise Safeguards](#-security-privacy--enterprise-safeguards)
11. [Roadmap & Future Vision](#️-roadmap--future-vision)
12. [Contributing](#-contributing)
13. [Acknowledgments & Credits](#-acknowledgments--credits)
14. [License](#-license)

---

## 🌟 What is Your Co-Pilot for Excel?

**Your Co-Pilot** is an advanced, free, and open-source alternative to **Microsoft 365 Copilot for Excel**. Running natively as a custom **Excel Add-in** in desktop and web versions of Excel, it equips your spreadsheets with cutting-edge artificial intelligence.

Unlike conventional chat assistants that merely generate code for you to copy and paste, **Your Co-Pilot is autonomous**: it reads your spreadsheet context, writes formulas, generates interactive charts, builds PivotTables, cleans messy datasets, applies conditional formatting, and executes arbitrary sandboxed Excel.js scripts in real time.

> **🔍 Search Keywords:** Open source Excel AI add-in, free alternative to Microsoft Copilot Excel, AI formula generator, automated PivotTable creator, Excel chart generator, ChatGPT in Excel, Claude in Excel, Gemini Excel add-in, local LLM Excel add-in, Ollama Excel integration.

> **🔒 Privacy First & Zero Telemetry:** API keys are stored strictly in your browser's local storage (`localStorage`). Your data is transmitted _only_ to the AI provider you configure — never through any third-party telemetry, proxy, or logging server.

---

## 💎 Why Choose Your Co-Pilot Over Paid Alternatives?

| Feature | Microsoft 365 Copilot ($30/mo) | **Your Co-Pilot (Open Source)** |
|---------|:-----------------------------:|:-----------------------------:|
| **Supported AI Providers** | OpenAI GPT-4 Only | **7 Providers** (NVIDIA NIM, Gemini, Groq, OpenRouter, OpenAI, Anthropic, Ollama) |
| **Offline / Air-Gapped Mode** | ❌ Cloud & Microsoft Account Required | ✅ **100% Offline & Private** (via Ollama local models) |
| **Custom Model Integration** | ❌ Locked to official endpoint | ✅ **Custom Model Manager** for any OpenAI-compatible API |
| **Excel Autonomous Actions** | Limited formulas & charts | ✅ **40+ Direct Actions** + sandboxed `eval_js` execution |
| **Formula Intelligence** | Basic formulas | ✅ Dynamic Arrays (`XLOOKUP`, `FILTER`, `UNIQUE`, `LET`, `LAMBDA`) |
| **Prompt Templates** | ❌ Limited | ✅ **50+ Curated Excel Templates** built-in |
| **Document Vision / Multimodal** | ❌ Text only | ✅ Attach screenshots/images for AI vision analysis |
| **Web Search Integration** | ❌ Disabled in taskpane | ✅ Live Web Search toggle built into navigation |
| **Office Version Support** | M365 Cloud Subscription Only | ✅ **Excel 2016, 2019, 2021, M365 Desktop & Excel Online** |
| **Price** | $30 / user / month | **$0 (100% Free & Open Source)** |

---

## ⚡ 40+ Excel Autonomous Actions & Sandboxing

Your Co-Pilot manipulates your workbook in real time using structured command blocks (`<EXCEL_CMD>`) emitted by the AI:

### 1. Formulas & Range Operations
- **Dynamic Array Formulas:** Automatically writes modern formulas (`XLOOKUP`, `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `LET`, `LAMBDA`).
- **2D Data Writing:** Writes full multi-row, multi-column arrays (`write_range`) with automatic data-type detection.
- **Cell & Range Manipulation:** `write_cell`, `clear_range`, `insert_range`, `delete_range`, and `merge_cells`.

### 2. Visualization & Charts
- **Chart Generation:** Creates Column, Bar, Line, Pie, Area, and Scatter charts (`create_chart`).
- **Chart Customization:** Configures titles, legend placement, and data labels (`format_chart`).
- **Inline Sparklines:** Injects trend sparklines directly into data cells (`add_sparklines`).

### 3. Data Analysis & Cleaning
- **PivotTables & Structured Tables:** Builds PivotTables with row/value fields and creates formatted Excel Tables (`create_table`, `create_pivot_table`).
- **Deduplication & Hygiene:** Removes duplicate rows (`remove_duplicates`), trims whitespace (`trim_whitespace`), standardizes casing (`change_case`), and strips empty rows (`remove_blank_rows`).
- **Sorting & Filtering:** Multi-column sorting (`sort_range`) and multi-criteria column filtering (`apply_filter`, `clear_filter`).

### 4. Formatting & Navigation
- **Conditional Formatting:** Color scales, top/bottom highlights, duplicate highlights, and data validation dropdowns.
- **Grid Layout:** Auto-fits columns and rows (`autofit_columns`, `autofit_rows`), freezes/unfreezes header panes (`freeze_panes`), and manages worksheet tabs (`add_sheet`, `delete_sheet`).
- **Named Ranges & CSV Extraction:** Creates/deletes named ranges and extracts ranges as token-efficient CSV (`get_range_csv`).

### 5. Sandboxed Dynamic Execution (`eval_js`)
- Inspired by Secure ECMAScript (SES) sandboxing, executes on-the-fly, arbitrary `Excel.run(async (context) => { ... })` scripts when standard actions do not cover a specialized workflow.

---

## 🔌 7 Supported AI Providers (Free & Paid)

Switch providers instantly from the **Settings Panel** without reloading Excel:

| Provider | Free Tier Available? | Recommended Models | Best For | Get API Key |
|----------|:-------------------:|-------------------|----------|-------------|
| **NVIDIA NIM** | ✅ **Generous Free Credits** | Llama 3.1 70B/405B, Mixtral, Nemotron | High-accuracy reasoning & formula writing | [build.nvidia.com](https://build.nvidia.com/) |
| **Google Gemini** | ✅ **Free (15 RPM)** | Gemini 1.5 Pro, Gemini 1.5 Flash | Massive spreadsheets (1M+ token context) | [aistudio.google.com](https://aistudio.google.com/) |
| **Groq** | ✅ **100% Free Tier** | Llama 3.1 70B, Llama 3 8B | Sub-second ultra-fast inference | [console.groq.com](https://console.groq.com/) |
| **OpenRouter** | ✅ **Free Models Available** | Open & proprietary model catalog | Testing diverse open-source LLMs | [openrouter.ai](https://openrouter.ai/) |
| **OpenAI** | 💰 Paid | GPT-4o, GPT-4o-mini | Industry-standard coding & math | [platform.openai.com](https://platform.openai.com/) |
| **Anthropic** | 💰 Paid | Claude 3.5 Sonnet, Claude 3 Haiku | In-depth financial analysis & explanation | [console.anthropic.com](https://console.anthropic.com/) |
| **Ollama** | ✅ **100% Free & Offline** | Llama 3.1, CodeLlama, DeepSeek | **Air-gapped, zero-egress data privacy** | [ollama.com](https://ollama.com/) |

> **💡 Zero-Cost Setup:** You can run this add-in completely free using NVIDIA NIM, Google Gemini, Groq, or local Ollama models.

---

## 🤖 Intelligent Chat & Context Experience

- **Real-Time Token Streaming:** Watch responses stream with smooth typewriter rendering and an instant **Stop/Abort** button (`AbortController`).
- **Live Spreadsheet Awareness:** Automatically reads your active cell selection, active sheet, all worksheet names, and named ranges to provide precise, context-aware answers.
- **Multimodal Image/Document Attachments:** Attach screenshots, invoices, or receipts via the paperclip button for AI vision analysis and automated data extraction into Excel.
- **50+ Curated Excel Prompt Templates:** Instant templates for Financial Modeling, Data Cleaning, Statistical Analysis, Dynamic Array Formulas, and VBA/Macro generation.
- **Multi-Conversation Management:** Search past conversations, switch threads, and persist chat history in local storage.
- **Web Search Integration:** Toggle live web search in the navigation bar for up-to-date market, currency, or industry data.

---

## 🚀 Step-by-Step Installation & Quick Start

### Option A: Deploy to Vercel (Recommended — 1-Click Free HTTPS Hosting)

Microsoft Office Add-ins require an `https://` URL. Vercel provides free HTTPS hosting in minutes:

1. **Fork or Clone this repository:**
   ```bash
   git clone https://github.com/SuprakashBakali/Your-Office-Copilot.git
   cd Your-Office-Copilot
   ```
2. **Deploy to Vercel:**
   - Go to [vercel.com](https://vercel.com) → **Add New Project** → Import your repository → Click **Deploy**.
   - Copy your deployed Vercel HTTPS domain URL (e.g., `https://your-office-copilot.vercel.app`).
3. **Update Manifest:**
   - Open `manifests/excel.xml`.
   - Replace all occurrences of `https://your-office-copilot.vercel.app` with your Vercel URL.
4. **Sideload in Microsoft Excel:**
   - In Excel (Desktop or Online): Go to **Insert → Add-ins → Upload My Add-in** (or **My Add-ins → Manage My Add-ins → Upload My Add-in**) and select `manifests/excel.xml`.

---

### Option B: Run Locally (Development Server)

For developers wanting to run and modify the add-in locally:

```bash
# 1. Clone the repository
git clone https://github.com/SuprakashBakali/Your-Office-Copilot.git
cd Your-Office-Copilot

# 2. Install Node.js dependencies
npm install

# 3. Start the Webpack Dev Server with HTTPS on port 3000
npm start

# 4. In manifests/excel.xml, update URLs to https://localhost:3000

# 5. In Excel, sideload manifests/excel.xml via Insert → Add-ins → Upload My Add-in
```

> **Note on Windows Dev Certificates:** If Excel warns about untrusted local HTTPS certificates, run `npx office-addin-dev-certs install` in your terminal.

---

### Option C: Shared Folder Sideloading (Desktop Excel 2016 / 2019 / 2021)

For older desktop versions of Excel without cloud add-in catalogs:

1. Create a shared folder on your Windows PC (e.g., `C:\MyAddins`).
2. Copy `manifests/excel.xml` into `C:\MyAddins`.
3. Right-click folder → **Properties → Sharing → Share** with your user account.
4. In Excel: Go to **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
5. Add the UNC path (e.g., `\\Your-PC\MyAddins`) and check **"Show in Menu"**.
6. Restart Excel → Click **Insert → My Add-ins → Shared Folder** → Select **Your Co-Pilot**.

---

## 🏗️ Technical Architecture & Project Structure

```
Your-Office-Copilot/
├── manifests/                    # Office Add-in XML Manifest
│   └── excel.xml                 # Dedicated Excel Add-in manifest
├── assets/                       # Custom Ribbon icons (16, 32, 64, 80, 128px)
├── src/
│   ├── commands/                 # Office Ribbon command handlers (showTaskpane)
│   │   ├── commands.html
│   │   └── commands.ts
│   └── taskpane/                 # React 18 Single-Page Application (SPA)
│       ├── components/
│       │   ├── chat/             # ChatPanel, ChatInput, MessageBubble
│       │   ├── layout/           # NavigationTabs, Sidebar layout
│       │   ├── settings/         # SettingsPanel, ApiKeyManager, CustomModelManager, ModelSelector
│       │   └── shared/           # MarkdownRenderer, CodeBlock, CopyButton, ThinkingBlock, Toast
│       ├── hooks/                # useChat, useAI, useTheme, useSettings
│       ├── services/
│       │   ├── ai/               # 7 Provider integrations + ProviderFactory + Types
│       │   ├── office/           # ExcelService.ts (40+ actions), OfficeContext.ts
│       │   ├── compaction.ts     # Conversation token optimization
│       │   └── prompts.ts        # Dynamic Excel system prompt builder
│       ├── store/                # AppContext React Context state store
│       ├── styles/               # Fluent UI v9 theme tokens + global CSS
│       ├── types/                # TypeScript interfaces (CellData, ExcelCommand, Message, etc.)
│       └── utils/                # Formatters, storage encryption, CSV utilities
│   ├── App.tsx
│   ├── index.tsx
│   └── taskpane.html
├── package.json
├── tsconfig.json
├── webpack.config.js
└── vercel.json
```

### Technology Stack

| Component | Technology |
| :--- | :--- |
| **Frontend Framework** | React 18.3 + TypeScript 5.5 |
| **Design System** | Microsoft Fluent UI v9 (`@fluentui/react-components`, `@fluentui/react-icons`) |
| **Excel Integration** | Official Microsoft Excel JavaScript API (`Office.js` / `Excel.run`) |
| **Bundler & Tooling** | Webpack 5 with code splitting, Hot Module Replacement (HMR) |
| **Markdown & Code** | `react-markdown` + `remark-gfm` + Prism syntax highlighting |
| **State Management** | React Context API + Custom Hooks (`useChat`, `useSettings`, `useTheme`) |

---

## 🛠️ Development & Validation Commands

```bash
# Start development server with HTTPS
npm start

# Run TypeScript type checker
npm run typecheck

# Lint source code
npm run lint

# Auto-fix lint issues
npm run lint:fix

# Format code with Prettier
npm run format

# Validate Excel manifest against Microsoft schema
npm run validate:excel

# Build optimized production bundle
npm run build
```

---

## 📖 Deep-Dive Documentation

### 1. Application Overview & Fluent UI v9 Design System

Your Co-Pilot runs as a high-performance **Single Page Application (SPA)** embedded directly inside the Excel Taskpane on Windows, macOS, and Web browsers.

```
+-----------------------------------------------------------------------+
|                         MICROSOFT EXCEL WINDOW                        |
|  +--------------------------------------------+  +-----------------+  |
|  |                                            |  |  TASKPANE       |  |
|  |             ACTIVE WORKBOOK                |  |  Your Co-Pilot  |  |
|  |                                            |  |  +-----------+  |  |
|  |  [A1: Sales Data]                          |  |  | Nav Tabs  |  |  |
|  |  [B1: Revenue]                             |  |  +-----------+  |  |
|  |  [C1: Formulas / Dynamic Arrays]           |  |  | ChatView  |  |  |
|  |                                            |  |  |    or     |  |  |
|  |                                            |  |  | Settings  |  |  |
|  |                                            |  |  +-----------+  |  |
|  |                                            |  |  | ChatInput |  |  |
|  +--------------------------------------------+  +-----------------+  |
+-----------------------------------------------------------------------+
```

- **Fluent UI v9 Integration:** Built with official Microsoft Fluent UI v9 design tokens for a native appearance matching Microsoft 365.
- **Adaptive Theme Modes:** Supports `Light`, `Dark`, and `Auto` modes (dynamically syncing with the host Excel theme).
- **Responsive Layout:** Fluid design adapted for narrow (300px) and wide sidebar taskpanes.

---

### 2. Excel Ribbon Integration

- **Home Tab Ribbon Button:** Injects a custom **Your Co-Pilot** button into the Excel Home tab via `manifests/excel.xml`.
- **Command Invocation:** Clicking the button executes `showTaskpane`, opening the AI assistant without reloading the workbook.
- **Multi-Resolution Icons:** Ships with `16x16`, `32x32`, `64x64`, `80x80`, and `128x128` icons for sharp rendering across standard and high-DPI displays.

---

### 3. Primary Navigation & Conversation Manager

The toolbar (`NavigationTabs.tsx`) sits at the top of the taskpane:

```
+-----------------------------------------------------------------------+
|  [💬 Chat]  |  [📄 Context: ON]  [🌐 Web]  [📜 History]  [➕ New]  [🗑]  |  [⚙️ Settings]  |
+-----------------------------------------------------------------------+
```

- **View Switcher:** Seamlessly toggles between Chat and Settings views.
- **Context Toggle:** Enables/disables live cell selection inclusion in prompts (green when active).
- **Web Search Toggle:** Toggles live web search for market and external information.
- **History Menu:** Searchable dropdown of past conversations with message counts and timestamps.
- **New / Delete Thread:** Quick actions to start fresh threads or delete the active conversation.

---

### 4. Chat Panel & Live Range Context

The conversation engine (`ChatPanel.tsx` & `ChatInput.tsx`) features:

- **Live Range Banner:** Dynamically queries `ExcelService.getSelectedRange()` to display the active selection (e.g., `Selected: A1:D50 on Sheet1`).
- **Attachment Support:** Paperclip button allows attaching screenshots or reference images for multimodal analysis.
- **Prompt Drawer:** Categorized library of 50+ Excel prompt templates (Formulas, Visualizations, Financial Analysis, Data Cleaning).
- **Action Execution Badges:** Displays execution pills for actions performed by the AI (e.g., `⚡ Created Chart 'Revenue Summary'`).
- **Code Block Copier:** Syntax-highlighted formula and code blocks with 1-click clipboard copy.

---

### 5. Settings & Multi-Provider API Configuration

The configuration hub (`SettingsPanel.tsx`):

- **7 Provider Selectors:** Choose between NVIDIA NIM, Google Gemini, Groq, OpenRouter, OpenAI, Anthropic, and Ollama.
- **Custom Model Manager:** Add private OpenAI-compatible endpoints with custom base URLs and headers.
- **API Key Storage:** Direct client-side `localStorage` storage with live connection testing.
- **Context Guardrails:** Configurable cell serialization caps (default: 500 cells) to optimize token usage.

---

### 6. Excel Automation Engine (`ExcelService.ts`)

| Category | Supported Operations | Description |
| :--- | :--- | :--- |
| **Writing & Formulas** | `write_cell`, `write_formula`, `write_range` | Writes values, dynamic arrays (`FILTER`, `XLOOKUP`, `UNIQUE`), and 2D matrices. |
| **Visualization** | `create_chart`, `delete_chart`, `format_chart`, `add_sparklines` | Column, line, bar, pie, scatter charts; inline cell sparklines. |
| **Data Analysis** | `create_pivot_table`, `create_table`, `sort_range`, `apply_filter`, `clear_filter` | PivotTables, structured tables, multi-column sorting and filtering. |
| **Formatting** | `format_range`, `add_conditional_formatting`, `highlight_duplicates`, `highlight_top_bottom` | Font colors, fills, borders, number formatting, color scales. |
| **Data Cleaning** | `remove_duplicates`, `trim_whitespace`, `change_case`, `remove_blank_rows` | Purges duplicate rows, trims excess spaces, cleans blank rows. |
| **Worksheet Layout** | `add_sheet`, `delete_sheet`, `insert_range`, `delete_range`, `merge_cells`, `group_data`, `ungroup_data` | Tab management, row/column block insertions, headers merging. |
| **Navigation & Panes** | `freeze_panes`, `unfreeze_panes`, `autofit_columns`, `autofit_rows` | Locks header panes and auto-fits dimensions to content. |
| **Named Ranges & CSV** | `create_named_range`, `delete_named_range`, `search_data`, `get_all_objects`, `get_range_csv` | Named references, regex search, token-efficient CSV extraction. |
| **Escape Hatch** | `eval_js` | **SES Sandboxed Execution** of arbitrary `Excel.run(async (context) => { ... })` scripts. |

---

### 7. Shared UI Components & Utilities

- **`CodeBlock.tsx`:** Prism-highlighted formula and script blocks with copy buttons.
- **`ThinkingBlock.tsx`:** Collapsible reasoning trace for reasoning-capable models (e.g., DeepSeek R1, Claude Thinking).
- **`MarkdownRenderer.tsx`:** Custom Fluent-styled markdown renderer supporting tables, bolding, and lists.
- **`NotificationToast.tsx`:** Transient feedback alerts for copy, save, and action execution status.

---

### 8. State Management & Data Flow

```
                  +-----------------------------------+
                  |        AppContext Provider        |
                  |     (View, Theme, Host: Excel)    |
                  +-----------------+-----------------+
                                    |
          +-------------------------+-------------------------+
          |                                                   |
          v                                                   v
+-------------------+                               +-------------------+
|    useChat.ts     |                               |  useSettings.ts   |
| - Conversations   |                               | - API Keys        |
| - Live Context    |                               | - Custom Models   |
| - Streaming State |                               | - Theme & Limits  |
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

---

## 🔐 Security, Privacy & Enterprise Safeguards

- **Zero Remote Telemetry:** We host no intermediary servers, proxies, or databases. The add-in connects directly from your browser/Excel WebView to the AI provider endpoint.
- **Local Key Storage:** API keys reside exclusively in browser `localStorage`.
- **Sandboxed Execution:** All workbook modifications execute strictly within Excel's official `Excel.run(async (context) => { ... })` `RequestContext` sandbox.
- **Air-Gapped Privacy:** For confidential financial spreadsheets, configure Ollama in Settings to use local LLMs (Llama 3.1 8B, DeepSeek) with zero internet egress.

---

## 🗺️ Roadmap & Future Vision

- [x] **Phase 1 (Core):** 7 AI Providers, Fluent UI v9 Taskpane, Excel automation engine, streaming chat.
- [x] **Phase 2 (Sandboxing):** Dynamic `eval_js` execution for unconstrained custom Excel workflows.
- [x] **Phase 3 (Vision):** Multimodal attachment support for screenshot and document data extraction.
- [ ] **Phase 4 (Custom Worksheet Functions):** Direct cell formulas (e.g., `=AI.ANALYZE(A1:D50, "Summarize revenue")`).
- [ ] **Phase 5 (Multi-Step Autonomous Agents):** Goal-oriented multi-step execution (e.g., *"Clean data, generate pivot table, and format financial dashboard"*).

---

## 🤝 Contributing

Contributions are warmly welcomed! Whether it's adding new prompt templates, optimizing Excel automation commands, or expanding provider support:

1. **Fork** the repository on GitHub.
2. **Create** your feature branch: `git checkout -b feature/amazing-feature`.
3. **Commit** your changes: `git commit -m 'Add amazing feature'`.
4. **Push** to the branch: `git push origin feature/amazing-feature`.
5. **Open** a Pull Request against `main`.

---

## 🙏 Acknowledgments & Credits

Special thanks to **[@hewliyang](https://github.com/hewliyang)** and their [office-agents](https://github.com/hewliyang/office-agents) repository. Their Secure ECMAScript (SES) sandboxing implementation inspired the dynamic `eval_js` execution engine in Your Co-Pilot.

---

## 📄 License

This project is open-source and released under the [MIT License](LICENSE). Free for personal, academic, and commercial enterprise use.

---

<p align="center">
  <strong>⭐ Star this repository if Your Co-Pilot helped you in Excel!</strong>
</p>

<p align="center">
  Built with ❤️ using React 18, TypeScript, Fluent UI v9, and Office.js
</p>
