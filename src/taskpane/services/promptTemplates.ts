/**
 * Prompt Templates — built-in, host-aware prompt library.
 *
 * Inspired by office-agents' Skills system (loaded SKILL.md files that the AI
 * reads on demand) and OfficeCLI's specialized skill triggers (pitch-deck,
 * financial-model, data-dashboard, academic-paper).
 *
 * Here we ship a curated set of one-click templates as data. Clicking a
 * template drops its prompt into the chat input so the user can hit Send.
 * This keeps the system lean (no VFS, no skill-install UI) while giving
 * users the same "I want to do X" affordance.
 */
import { OfficeHostType } from '../types';

export interface PromptTemplate {
  id: string;
  title: string;
  description: string;
  /** Which Office hosts this template applies to. Empty = all. */
  hosts?: OfficeHostType[];
  /** Short emoji or icon hint for the UI. */
  icon: string;
  /** The prompt body. Use {placeholder} for user-fillable fields. */
  prompt: string;
  /** Optional category for grouping in the UI. */
  category: 'analyze' | 'create' | 'format' | 'explain' | 'utility';
}

export const PROMPT_TEMPLATES: PromptTemplate[] = [
  // ---- Excel ----
  {
    id: 'excel-summarize',
    title: 'Summarize Data',
    description: 'Get a statistical summary of the selected range',
    hosts: ['Excel'],
    icon: '📊',
    category: 'analyze',
    prompt: 'Analyze the selected range and give me a statistical summary: row/column count, data types, missing values, min/max/mean/median for numeric columns, unique value counts for text columns, and 3 notable patterns or outliers.',
  },
  {
    id: 'excel-dashboard',
    title: 'Build Dashboard',
    description: 'Create a KPI dashboard with charts from the current data',
    hosts: ['Excel'],
    icon: '📈',
    category: 'create',
    prompt: 'Build a dashboard on a new sheet named "Dashboard". From the data on the active sheet, create: (1) 3 KPI cards at the top showing key metrics, (2) a column chart of the top 10 items by the primary numeric column, (3) a pie chart of category distribution, (4) a line chart of any time-series column. Apply conditional formatting to highlight values above average. Auto-fit all columns when done.',
  },
  {
    id: 'excel-financial-model',
    title: 'Financial Model',
    description: 'Build a 3-statement model with assumptions, P&L, BS, CF',
    hosts: ['Excel'],
    icon: '💰',
    category: 'create',
    prompt: 'Build a 3-statement financial model on this workbook with the following sheets: (1) "Assumptions" — revenue growth %, gross margin %, opex growth %, tax rate %, all color-coded blue for inputs, (2) "P&L" — 5-year projection with revenue, COGS, gross profit, opex, EBITDA, depreciation, EBIT, tax, net income, (3) "Balance Sheet" — 5-year with current assets, PP&E, current liabilities, equity, (4) "Cash Flow" — indirect method linking P&L and BS changes. Use formulas everywhere except blue input cells. Format numbers as $#,##0;($#,##0) and percentages as 0.0%.',
  },
  {
    id: 'excel-clean-data',
    title: 'Clean Data',
    description: 'Remove duplicates, trim whitespace, fix casing',
    hosts: ['Excel'],
    icon: '🧹',
    category: 'utility',
    prompt: 'Clean the data on the active sheet: (1) trim leading/trailing whitespace from all text cells, (2) remove fully-blank rows, (3) highlight duplicate rows in pink, (4) standardize date columns to ISO format, (5) detect and report any mixed-type columns. Show me a summary of what you changed before applying.',
  },
  {
    id: 'excel-formula-explain',
    title: 'Explain Formula',
    description: 'Explain the formula in the selected cell in plain English',
    hosts: ['Excel'],
    icon: '🔍',
    category: 'explain',
    prompt: 'Look at the formula in my currently selected cell. Explain in plain English what it does, step by step. Then suggest 1-2 modern alternatives (e.g. XLOOKUP instead of VLOOKUP, dynamic arrays, LET/LAMBDA) if applicable.',
  },
  {
    id: 'excel-pivot-table',
    title: 'Auto Pivot Table',
    description: 'Create a pivot table summarizing the data',
    hosts: ['Excel'],
    icon: '🔄',
    category: 'create',
    prompt: 'Create a pivot table from the data on the active sheet. Choose the most logical row field, value field (sum), and column field based on the data. Place it on a new sheet named "Pivot". Format the values with thousand separators.',
  },

  // ---- Word ----
  {
    id: 'word-executive-summary',
    title: 'Executive Summary',
    description: 'Extract a 1-page executive summary from the document',
    hosts: ['Word'],
    icon: '📝',
    category: 'analyze',
    prompt: 'Read the current document and write a 1-page executive summary at the end of the document. Include: (1) the document\'s main thesis in 2 sentences, (2) 3-5 key findings as bullet points, (3) the recommended next steps. Apply Heading 1 style to "Executive Summary" and Heading 2 to subsections.',
  },
  {
    id: 'word-proofread',
    title: 'Proofread & Fix',
    description: 'Fix grammar, spelling, and clarity issues',
    hosts: ['Word'],
    icon: '✅',
    category: 'utility',
    prompt: 'Proofread the current document. Highlight spelling errors in yellow, grammar issues in pink, and style/clarity issues in light blue. Do NOT auto-correct — just highlight them and add a comment-style note after each issue explaining the suggested fix. At the end, insert a paragraph with a count of issues found per category.',
  },
  {
    id: 'word-academic-paper',
    title: 'Academic Paper',
    description: 'Format as APA-style academic paper with sections',
    hosts: ['Word'],
    icon: '🎓',
    category: 'create',
    prompt: 'Reformat the current document as an APA-style academic paper: (1) Title page with title, author, institution, (2) Abstract page (150-250 words), (3) Body with Introduction, Methods, Results, Discussion, Conclusion as Heading 1, (4) References section. Apply 1-inch margins implicitly (acknowledge if you can\'t), double-spacing (set line spacing), and Times New Roman 12pt where applicable.',
  },
  {
    id: 'word-rewrite-tone',
    title: 'Rewrite Tone',
    description: 'Rewrite the document in a different tone',
    hosts: ['Word'],
    icon: '🎭',
    category: 'format',
    prompt: 'Ask me which tone I want (formal / casual / persuasive / technical / simple), then rewrite each paragraph of the current document in that tone. Preserve all factual content, headings, and structure. Insert the rewritten version below each original paragraph, separated by a horizontal rule, so I can compare.',
  },
  {
    id: 'word-table-of-contents',
    title: 'Generate TOC',
    description: 'Insert a table of contents based on headings',
    hosts: ['Word'],
    icon: '📚',
    category: 'utility',
    prompt: 'Insert a Table of Contents at the top of the document based on the existing Heading 1, Heading 2, and Heading 3 styles. Format it as "Table of Contents" heading + a numbered list of sections with page numbers (use Word\'s built-in TOC if possible via eval_js, otherwise insert a manual list).',
  },

  // ---- PowerPoint ----
  {
    id: 'ppt-pitch-deck',
    title: 'Investor Pitch Deck',
    description: 'Generate a 10-slide investor pitch deck outline',
    hosts: ['PowerPoint'],
    icon: '💼',
    category: 'create',
    prompt: 'Create a 10-slide investor pitch deck. Slides: (1) Title slide with company name placeholder, (2) Problem, (3) Solution, (4) Market Size, (5) Product Demo, (6) Business Model, (7) Traction, (8) Team, (9) Financials, (10) Ask. For each slide, add a title textbox and a body textbox with 3-5 bullet points of placeholder content. Set speaker notes on each slide summarizing what to say.',
  },
  {
    id: 'ppt-from-doc',
    title: 'Slides from Document',
    description: 'Convert the current Word doc into a slide deck',
    hosts: ['PowerPoint'],
    icon: '🔄',
    category: 'create',
    prompt: 'I have a Word document open in another window. Ask me to paste its content (or summarize what it\'s about), then generate a 5-8 slide PowerPoint presentation from it: one slide per major section, with the section heading as the slide title and 3-4 bullet points from the section content. Add speaker notes that summarize each section in 1-2 sentences.',
  },
  {
    id: 'ppt-speaker-notes',
    title: 'Speaker Notes',
    description: 'Generate speaker notes for all slides',
    hosts: ['PowerPoint'],
    icon: '🗣️',
    category: 'create',
    prompt: 'Generate speaker notes for every slide in the current presentation. For each slide, the notes should: (1) summarize the slide\'s main point in 1 sentence, (2) provide 2-3 talking points expanding on the bullets, (3) suggest a transition phrase to the next slide. Keep each note under 100 words.',
  },
  {
    id: 'ppt-design-polish',
    title: 'Design Polish',
    description: 'Apply consistent colors and fonts across slides',
    hosts: ['PowerPoint'],
    icon: '🎨',
    category: 'format',
    prompt: 'Apply consistent design across all slides in the presentation: (1) format all title text boxes with the same font (Calibri 32pt bold) and color (dark navy #1A2B4A), (2) format all body text boxes with Calibri 18pt in #333333, (3) add a thin accent bar shape at the bottom of each slide in #76B900. Report what you changed.',
  },

  // ---- Cross-host ----
  {
    id: 'explain-selection',
    title: 'Explain Selection',
    description: 'Explain the currently selected content',
    icon: '💡',
    category: 'explain',
    prompt: 'Look at my current selection and explain it to me. If it\'s a formula, explain what it does. If it\'s data, summarize the key insight. If it\'s text, identify the main idea and any issues. Be concise — 3-5 sentences max.',
  },
  {
    id: 'translate',
    title: 'Translate',
    description: 'Translate the selection to another language',
    icon: '🌐',
    category: 'utility',
    prompt: 'Ask me which language to translate to, then translate the selected content to that language. Preserve formatting where possible. Insert the translation below the original, separated by a horizontal rule, so I can compare side-by-side.',
  },
  {
    id: 'web-research',
    title: 'Web Research',
    description: 'Research a topic on the web and summarize',
    icon: '🌍',
    category: 'analyze',
    prompt: 'I want to research a topic. Ask me what topic, then enable web search and find 3-5 authoritative sources. Summarize each source in 2-3 sentences with the URL, then provide a synthesized answer to my original question with citations.',
  },
];

/** Filter templates by host (or return all if `host` is Unknown). */
export function getTemplatesForHost(host: OfficeHostType | 'Unknown'): PromptTemplate[] {
  if (host === 'Unknown') return PROMPT_TEMPLATES;
  return PROMPT_TEMPLATES.filter(t => !t.hosts || t.hosts.length === 0 || t.hosts.includes(host));
}

/** Group templates by category for UI display. */
export function groupTemplatesByCategory(templates: PromptTemplate[]): Record<string, PromptTemplate[]> {
  const groups: Record<string, PromptTemplate[]> = {};
  for (const t of templates) {
    if (!groups[t.category]) groups[t.category] = [];
    groups[t.category].push(t);
  }
  return groups;
}

/** Replace {placeholder} tokens in a prompt with values from a map. */
export function fillPromptTemplate(prompt: string, vars: Record<string, string>): string {
  return prompt.replace(/\{(\w+)\}/g, (_, key) => vars[key] ?? `{${key}}`);
}
