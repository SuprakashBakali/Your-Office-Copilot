<p align="center">
  <img src="assets/icon.jpg" alt="Your Co-Pilot Logo" width="128" height="128" style="border-radius: 24px; box-shadow: 0 10px 25px rgba(0,0,0,0.2);" />
</p>

<h1 align="center">Your Co-Pilot — Open-Source Microsoft Office AI Add-in</h1>

<p align="center">
  <strong>The Ultimate AI-Powered Copilot for Excel, Word & PowerPoint — Powered by 7 AI Providers & Autonomous Document Manipulation</strong>
</p>

<p align="center">
  <a href="#-quick-start"><img src="https://img.shields.io/badge/Quick_Start-↓-76B900?style=for-the-badge&logo=rocket" alt="Quick Start" /></a>
  <a href="PAGE_DETAILS.md"><img src="https://img.shields.io/badge/Page_Details-Documentation-blueviolet?style=for-the-badge&logo=read-the-docs" alt="Page Details" /></a>
  <a href="OPTIMIZATIONS.md"><img src="https://img.shields.io/badge/Performance-Optimizations-FFA500?style=for-the-badge&logo=speedtest" alt="Optimizations" /></a>
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-blue?style=for-the-badge" alt="MIT License" /></a>
  <img src="https://img.shields.io/badge/React-18-61DAFB?style=for-the-badge&logo=react&logoColor=white" alt="React 18" />
  <img src="https://img.shields.io/badge/TypeScript-5.5-3178C6?style=for-the-badge&logo=typescript&logoColor=white" alt="TypeScript 5.5" />
  <img src="https://img.shields.io/badge/Office.js-Add--in-D83B01?style=for-the-badge&logo=microsoftoffice&logoColor=white" alt="Office.js Add-in" />
</p>



---

## 🌟 What is Your Co-Pilot?

**Your Co-Pilot** is an advanced, open-source, and free alternative to **Microsoft 365 Copilot** and expensive enterprise AI add-ins. Running natively as a custom **Office Add-in** inside Excel, Word, and PowerPoint, it empowers your desktop and web Office applications with cutting-edge artificial intelligence.

Unlike traditional AI chat assistants that only give you text to copy and paste, **Your Co-Pilot is autonomous**: it can directly read your selection, write formulas, create charts, build PivotTables, insert formatted Word tables, and construct PowerPoint slides automatically in the background.

> **🔍 Search Keywords:** Open source AI office add-in, free alternative to Microsoft Copilot, AI for Excel, AI for Word, AI for PowerPoint, integrate ChatGPT in Excel, Claude in Word, Gemini for Office 365, local LLM Excel add-in, Ollama Office integration.

> **🔒 Privacy First & Zero Telemetry:** Your API keys are stored locally in your browser (`localStorage`). Your data is sent _only_ to the AI provider you explicitly configure — never to any third-party server or analytics tracking service.

---

## 💎 Why Choose Your Co-Pilot Over Enterprise Alternatives?

| Feature | Microsoft 365 Copilot / Paid Add-ins | **Your Co-Pilot (Open Source)** |
|---------|:-----------------:|:-----------------:|
| **Offline / Air-Gapped Mode** | ❌ Requires Cloud & Microsoft Account | ✅ **100% Offline Capable** (via local Ollama models) |
| **Custom Model Integration** | ❌ Locked to official endpoints | ✅ **Custom Model Manager** for any OpenAI-compatible API |
| **Excel Autonomous Actions** | ✅ Formulas & charts | ✅ **40+ direct commands** + sandboxed JavaScript `eval_js` |
| **Word Table & Style Automation** | ✅ Basic formatting | ✅ **Full table insertion**, style application & structure extraction |
| **PowerPoint Shape & Note Control** | ✅ Layout suggestions | ✅ **Direct shape/textbox injection** & automated speaker notes |
| **Prompt Library** | ❌ Limited prompts | ✅ **50+ Curated Templates** built-in across all Office hosts |
| **Office Version Support** | ❌ Microsoft 365 Cloud Subscription Only | ✅ **Excel/Word/PowerPoint 2016, 2019, 2021, M365 & Web** |

---

## ✨ Comprehensive Features & Capabilities

### ⚡ Automated Copilot Actions & Dynamic Sandboxing
Your Co-Pilot doesn't just suggest code—it **directly manipulates your documents** in real time using structured XML/JSON command blocks emitted by the AI:
- **📊 Excel Automation (40+ Actions):**
  - **Formulas & Ranges:** Automatically write dynamic array formulas (`XLOOKUP`, `FILTER`, `SORT`, `UNIQUE`), write 2D data arrays, and merge cells.
  - **Data Visualization:** Create Column, Pie, Line, Bar, Area, and Scatter charts; add inline Sparklines; configure legend positions and titles.
  - **Data Analysis & Cleaning:** Generate PivotTables, insert structured Excel tables, sort multi-column ranges, remove duplicates, trim whitespace, remove blank rows, and apply multi-criteria filters.
  - **Formatting:** Apply conditional formatting (color scales, top/bottom highlights, duplicate highlights), bold/color headers, autofit rows/columns, and freeze panes.
  - **Advanced Execution (`eval_js`):** Inspired by Secure ECMAScript (SES) sandboxing, executes on-the-fly, arbitrary `Excel.run(async (context) => { ... })` scripts when standard actions don't cover a niche workflow.
- **📝 Word Automation:**
  - Insert paragraphs at any location (`before`, `after`, `start`, `end`), generate multi-row formatted tables, apply built-in Word styles (`Heading 1`, `Heading 2`), clear messy formatting, perform mass find-and-replace, and highlight search terms.
- **🎯 PowerPoint Automation:**
  - Add new slides, inject positioned text boxes, draw geometric shapes (`Rectangle`, `Oval`, `Line`), format shape fills/fonts, and write comprehensive speaker notes.

### 🤖 Intelligent Chat & Streaming Experience
- **Real-time Token Streaming:** Watch responses stream word-by-word with smooth typewriter animations and an instant **Stop/Abort** button.
- **Context-Aware Selections:** Automatically inspects selected Excel cells, Word paragraphs, or PowerPoint slides (`includeContext` toggle) and passes structured `CellData` / `DocumentData` to the LLM.
- **Rich Markdown & Syntax Highlighting:** Powered by `react-markdown` and Prism syntax highlighting with a **1-Click Copy** button for code snippets.
- **Multi-Conversation Management:** Create new threads, switch between past conversations, or delete history from a persistent dropdown menu.

---

## 🔌 7 Supported AI Providers (Free & Paid)

Switch between providers instantly from the **Settings Panel** without restarting Excel, Word, or PowerPoint:

| Provider | Free Tier Available? | Best Models | Best For | Get API Key |
|----------|:---------:|-------------|----------|-------------|
| **NVIDIA NIM** | ✅ **Generous Free Credits** | Llama 3.1 70B/405B, Mixtral, Nemotron | High-accuracy reasoning & coding | [build.nvidia.com](https://build.nvidia.com/) |
| **Google Gemini** | ✅ **Free (15 req/min)** | Gemini 1.5 Pro, Gemini 1.5 Flash | Massive spreadsheets (1M+ context) | [aistudio.google.com](https://aistudio.google.com/) |
| **Groq** | ✅ **100% Free Tier** | Llama 3.1 70B, Llama 3 8B | Sub-second, ultra-fast responses | [console.groq.com](https://console.groq.com/) |
| **OpenRouter** | ✅ **Free models available** | Hundreds of open & proprietary models | Accessing any open-source model | [openrouter.ai](https://openrouter.ai/) |
| **OpenAI** | 💰 Paid | GPT-4o, GPT-4o-mini, GPT-3.5 Turbo | Industry-standard reasoning | [platform.openai.com](https://platform.openai.com/) |
| **Anthropic** | 💰 Paid | Claude 3.5 Sonnet, Claude 3 Haiku | Best creative writing & Word editing | [console.anthropic.com](https://console.anthropic.com/) |
| **Ollama** | ✅ **100% Free & Offline** | Llama 3.1, Mistral, CodeLlama, DeepSeek | **Air-gapped enterprise privacy** | [ollama.com](https://ollama.com/) |

> **💡 Zero-Cost Tip:** You can run this entire project for **$0** using NVIDIA NIM, Google Gemini, Groq, or offline Ollama models.


---

## 🚀 Step-by-Step Installation & Quick Start

### Option A: Deploy to Vercel (Recommended — 1-Click Free Hosting)

Microsoft Office Add-ins require an `https://` URL. Vercel provides free HTTPS hosting in seconds:

1. **Fork or Clone this repository:**
   ```bash
   git clone https://github.com/SuprakashBakali/office-ai-copilot.git
   cd office-ai-copilot
   ```
2. **Deploy to Vercel:**
   - Go to [vercel.com](https://vercel.com) → **Add New Project** → Import your repository → Click **Deploy**.
   - Copy your deployed Vercel domain URL (e.g., `https://office-ai-copilot.vercel.app`).
3. **Update Manifests:**
   - Open `manifests/excel.xml`, `manifests/word.xml`, and `manifests/powerpoint.xml`.
   - Replace all occurrences of `https://your-office-copilot.vercel.app` with your new Vercel HTTPS URL.
4. **Sideload in Microsoft Office:**
   - In Excel, Word, or PowerPoint: Go to **Insert → Add-ins → Upload My Add-in** (or **My Add-ins → Manage My Add-ins → Upload My Add-in**) and select your updated XML manifest file.

---

### Option B: Run Locally (Development Server)

For developers wanting to hack on the codebase locally:

```bash
# 1. Clone the repository
git clone https://github.com/SuprakashBakali/office-ai-copilot.git
cd office-ai-copilot

# 2. Install Node.js dependencies
npm install

# 3. Start the Webpack Dev Server with self-signed HTTPS certificates on port 3000
npm start

# 4. Update local manifest URLs if needed (for local dev, use https://localhost:3000 in manifests):
# Open manifests/excel.xml and replace https://your-office-copilot.vercel.app with https://localhost:3000

# 5. In Excel / Word / PowerPoint, sideload the local manifest:
# Insert → Add-ins → Upload My Add-in → manifests/excel.xml
```

> **Note on Windows Dev Certificates:** If Excel warns about untrusted local HTTPS certificates, run `npx office-addin-dev-certs install` once in your terminal.

---

### Option C: Shared Folder Sideloading (Desktop Office 2016 / 2019 / 2021)

For older Desktop versions of Office without Microsoft 365 cloud catalogs:

1. Create a shared network folder on your Windows computer (e.g., `C:\MyAddins`).
2. Copy `manifests/excel.xml` (or `word.xml`, `powerpoint.xml`) into `C:\MyAddins`.
3. Right-click folder → **Properties → Sharing → Share** with yourself.
4. In Excel: Go to **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
5. Add the network UNC path (e.g., `\\Your-PC\MyAddins`) and check **"Show in Menu"**.
6. Restart Excel → Click **Insert → My Add-ins → Shared Folder** → Select **Your Co-Pilot**.

---

## 🏗️ Technical Architecture & Project Structure

```
office-ai-copilot/
├── manifests/                    # Office Add-in XML Manifests (Excel, Word, PowerPoint)
│   ├── excel.xml
│   ├── word.xml
│   └── powerpoint.xml
├── assets/                       # Custom Ribbon icons & preview badges
├── src/
│   ├── commands/                 # Office Ribbon command actions (showTaskpane)
│   ├── taskpane/                 # React 18 Single-Page Application (SPA)
│   │   ├── components/
│   │   │   ├── chat/             # ChatPanel, MessageBubble, ChatInput
│   │   │   ├── excel/            # Excel specific tools & prompt categories
│   │   │   ├── word/             # Word specific tools & prompt categories
│   │   │   ├── powerpoint/       # PowerPoint specific tools & prompt categories
│   │   │   ├── settings/         # SettingsPanel, ApiKeyManager, CustomModelManager
│   │   │   ├── layout/           # NavigationTabs, Sidebar layout
│   │   │   └── shared/           # MarkdownRenderer, CodeBlock, CopyButton, Toast
│   │   ├── hooks/                # useAI, useChat, useTheme, useSettings
│   │   ├── services/
│   │   │   ├── ai/               # 7 AI Provider classes + ProviderFactory + Prompts
│   │   │   └── office/           # ExcelService, WordService, PowerPointService
│   │   ├── store/                # AppContext React Context state store
│   │   ├── styles/               # Fluent UI v9 tokens + global CSS
│   │   ├── types/                # TypeScript interfaces (CellData, DocumentData, etc.)
│   │   └── utils/                # localStorage encryption, CSV formatters
│   ├── App.tsx
│   └── index.tsx
├── PAGE_DETAILS.md               # Exhaustive UI & page architectural documentation
├── OPTIMIZATIONS.md              # Performance & token optimization reference
├── package.json
├── tsconfig.json
└── webpack.config.js
```

### Technology Stack

| Component | Technology Used |
| :--- | :--- |
| **Frontend Framework** | React 18 + TypeScript 5.5 |
| **Design System** | Microsoft Fluent UI v9 (`@fluentui/react-components`) |
| **Office Integration** | Official Microsoft Office.js Add-in API |
| **Bundler & Tooling** | Webpack 5 with Hot Module Replacement (HMR) |
| **Markdown Rendering** | `react-markdown` + `remark-gfm` + Prism syntax highlighting |
| **State Management** | React Context API + Custom Hooks (`useChat`, `useSettings`) |

---

## 🛠️ Development & Validation Commands

```bash
# Run TypeScript type checker across all source files
npm run typecheck

# Lint all TypeScript/TSX source code
npm run lint

# Auto-fix linting issues
npm run lint:fix

# Format source files with Prettier
npm run format

# Validate Office Add-in XML manifests against official schemas
npm run validate:excel
npm run validate:word
npm run validate:powerpoint

# Build optimized production bundle to /dist
npm run build
```

---

## 🔐 Security, Privacy & Enterprise Safeguards

- **Zero Data Retention:** We do not host any server, proxy, or logging database. Your Co-Pilot connects directly from your browser to your selected AI provider's API endpoint.
- **Client-Side Key Encryption:** API keys are stored only in your browser's local storage (`localStorage`).
- **Sandboxed Office APIs:** All document modifications execute strictly inside Office.js `RequestContext` sandboxes (`Excel.run`, `Word.run`, `PowerPoint.run`), preventing any unauthorized access to your operating system or local file system.
- **Air-Gapped Privacy:** For confidential financial or medical spreadsheets, configure Ollama in the Settings panel to run local LLMs (such as Llama 3.1 8B or CodeLlama) with zero internet egress.

---

## 🤝 Contributing

We love contributions from the open-source community! Whether it's adding new prompt templates, optimizing Excel automation commands, or supporting new AI models, here's how to contribute:

1. **Fork** the repository on GitHub.
2. **Create** your feature branch: `git checkout -b feature/amazing-feature`.
3. **Commit** your changes: `git commit -m 'Add amazing feature'`.
4. **Push** to the branch: `git push origin feature/amazing-feature`.
5. **Open** a Pull Request against `main`.

---

## 🙏 Acknowledgments & Credits

A special thanks to **[@hewliyang](https://github.com/hewliyang)** and their [office-agents](https://github.com/hewliyang/office-agents) repository. Their innovative implementation of Secure ECMAScript (SES) sandboxing heavily inspired the dynamic `evaluate_office_js` / `eval_js` automation tool in Your Co-Pilot, enabling secure, on-the-fly execution of arbitrary Office.js scripts!

---

## 📄 License

This project is open-source and released under the [MIT License](LICENSE). Free for personal, academic, and commercial enterprise use.

---

<p align="center">
  <strong>⭐ Star this repository if Your Co-Pilot saved you time in Microsoft Office!</strong>
</p>

<p align="center">
  Built with ❤️ using React 18, TypeScript, Fluent UI v9, and Office.js
</p>

<p align="center">
  Copyright (c) 2026 Suprakash Bakali
</p>
