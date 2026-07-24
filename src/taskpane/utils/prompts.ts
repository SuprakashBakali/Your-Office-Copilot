// ============================================================
// Prompt Templates — Pre-built prompts for each Office app
// ============================================================

import { PromptTemplate, OfficeHostType } from "../types";

export const PROMPT_TEMPLATES: PromptTemplate[] = [
  // ---- Excel: Analysis ----
  {
    id: "excel-analyze-selection",
    title: "Analyze Selected Data",
    description: "Get a statistical summary of the selected cells",
    prompt: "Analyze the selected data. Provide: count, sum, average, min, max, median, standard deviation, and any notable patterns or outliers. Format as a clear summary.",
    category: "analysis",
    hostApp: "Excel",
    icon: "📊",
  },
  {
    id: "excel-trend-analysis",
    title: "Trend Analysis",
    description: "Identify trends and patterns in your data",
    prompt: "Analyze the trends in this data. Identify: direction (increasing/decreasing/stable), rate of change, seasonal patterns, anomalies, and provide forecasting suggestions.",
    category: "analysis",
    hostApp: "Excel",
    icon: "📈",
  },
  {
    id: "excel-outlier-detection",
    title: "Detect Outliers",
    description: "Find outliers and anomalies",
    prompt: "Examine this data for outliers and anomalies. Use statistical methods (IQR, Z-score) to identify unusual values. Explain why each is flagged and suggest how to handle them.",
    category: "analysis",
    hostApp: "Excel",
    icon: "🔍",
  },
  {
    id: "excel-correlation",
    title: "Correlation Analysis",
    description: "Find relationships between columns",
    prompt: "Analyze the correlations between the columns in this data. Identify which variables are positively/negatively correlated, the strength of each relationship, and any surprising findings.",
    category: "analysis",
    hostApp: "Excel",
    icon: "🔗",
  },
  {
    id: "excel-duplicates",
    title: "Find Duplicates",
    description: "Detect duplicate rows or values",
    prompt: "Scan this data for duplicate values and rows. Report: which values appear more than once, their locations, and suggest a formula or approach to handle them.",
    category: "cleaning",
    hostApp: "Excel",
    icon: "🔄",
  },

  // ---- Excel: Formulas ----
  {
    id: "excel-explain-formula",
    title: "Explain Formula",
    description: "Break down a formula step by step",
    prompt: "Explain this Excel formula step by step in simple language. Break down each function, what it does, and how the parts work together. Also suggest if there's a simpler alternative.",
    category: "formula",
    hostApp: "Excel",
    icon: "🧮",
  },
  {
    id: "excel-generate-formula",
    title: "Generate Formula",
    description: "Create a formula from plain English",
    prompt: "Generate an Excel formula that does the following: [describe what you need]. Provide the formula, explain how it works, and give an example of how to use it.",
    category: "formula",
    hostApp: "Excel",
    icon: "✨",
  },
  {
    id: "excel-fix-formula",
    title: "Fix Formula Error",
    description: "Debug and fix formula errors",
    prompt: "This formula is producing an error. Diagnose the issue, explain why it's failing, and provide the corrected formula. Common issues to check: circular references, wrong range, missing parentheses, data type mismatches.",
    category: "formula",
    hostApp: "Excel",
    icon: "🔧",
  },
  {
    id: "excel-optimize-formula",
    title: "Optimize Formula",
    description: "Make a formula more efficient",
    prompt: "Review this formula and suggest optimizations. Consider: performance (avoid VLOOKUP in favor of INDEX/MATCH or XLOOKUP), readability, error handling (IFERROR), and modern Excel functions (FILTER, UNIQUE, SORT, LET).",
    category: "formula",
    hostApp: "Excel",
    icon: "⚡",
  },

  // ---- Excel: Charts ----
  {
    id: "excel-chart-recommend",
    title: "Recommend Chart",
    description: "Get the best chart type for your data",
    prompt: "Based on this data, recommend the best chart type(s) to visualize it. Consider: the data structure, what story it tells, and the audience. Suggest specific chart types with explanations.",
    category: "chart",
    hostApp: "Excel",
    icon: "📉",
  },

  // ---- Excel: Finance ----
  {
    id: "excel-budget-analysis",
    title: "Budget Analysis",
    description: "Analyze budget vs actual spending",
    prompt: "Analyze this budget data. Calculate: variance (amount and percentage), identify over/under-budget categories, suggest areas for cost reduction, and provide a summary dashboard formula set.",
    category: "finance",
    hostApp: "Excel",
    icon: "💰",
  },
  {
    id: "excel-financial-ratios",
    title: "Financial Ratios",
    description: "Calculate key financial ratios",
    prompt: "From this financial data, calculate key ratios: liquidity (current ratio, quick ratio), profitability (gross margin, net margin, ROE, ROA), efficiency (asset turnover, inventory turnover), and leverage (debt-to-equity). Provide formulas and interpretation.",
    category: "finance",
    hostApp: "Excel",
    icon: "📋",
  },
  {
    id: "excel-cashflow",
    title: "Cash Flow Analysis",
    description: "Analyze cash flow patterns",
    prompt: "Analyze the cash flow data. Identify: inflow/outflow patterns, net cash position trends, seasonal variations, and provide projections. Suggest formulas for a cash flow statement.",
    category: "finance",
    hostApp: "Excel",
    icon: "💵",
  },
  {
    id: "excel-breakeven",
    title: "Break-Even Analysis",
    description: "Calculate break-even point",
    prompt: "Using this data, calculate the break-even point. Consider: fixed costs, variable costs per unit, and selling price per unit. Provide the break-even formula in units and revenue, plus a sensitivity analysis.",
    category: "finance",
    hostApp: "Excel",
    icon: "⚖️",
  },

  // ---- Excel: Data Cleaning ----
  {
    id: "excel-clean-data",
    title: "Clean This Data",
    description: "AI-assisted data cleaning",
    prompt: "Review this data for quality issues. Check for: inconsistent formatting, mixed data types, extra spaces, special characters, missing values, typos, and inconsistent naming. Suggest specific formulas to fix each issue.",
    category: "cleaning",
    hostApp: "Excel",
    icon: "🧹",
  },
  {
    id: "excel-standardize-dates",
    title: "Standardize Dates",
    description: "Fix inconsistent date formats",
    prompt: "These cells contain dates in various formats. Identify all the different formats present and provide a formula to convert them all to a standard format (YYYY-MM-DD). Handle edge cases like text dates.",
    category: "cleaning",
    hostApp: "Excel",
    icon: "📅",
  },

  // ---- Excel: Code Generation ----
  {
    id: "excel-vba-macro",
    title: "Generate VBA Macro",
    description: "Create a VBA macro for automation",
    prompt: "Generate a VBA macro that does the following: [describe the task]. Include: proper error handling, comments explaining each step, and instructions on how to add it to the workbook (Alt+F11).",
    category: "code",
    hostApp: "Excel",
    icon: "⚙️",
  },
  {
    id: "excel-office-script",
    title: "Generate Office Script",
    description: "Create an Office Script (TypeScript)",
    prompt: "Generate an Office Script (TypeScript) that does the following: [describe the task]. Include comments and explain how to run it from the Automate tab in Excel.",
    category: "code",
    hostApp: "Excel",
    icon: "📜",
  },
  {
    id: "excel-power-query",
    title: "Generate Power Query (M)",
    description: "Create Power Query M code",
    prompt: "Generate Power Query M code that does the following: [describe the data transformation]. Include step-by-step comments and instructions on how to paste it into the Advanced Editor.",
    category: "code",
    hostApp: "Excel",
    icon: "🔌",
  },
  {
    id: "excel-sql",
    title: "Generate SQL Query",
    description: "Create SQL from your data structure",
    prompt: "Based on this data structure, generate SQL queries for: creating the table, inserting the data, and common queries (SELECT, GROUP BY, JOIN, aggregations). Use standard SQL syntax.",
    category: "code",
    hostApp: "Excel",
    icon: "🗃️",
  },

  // ---- Word: Writing ----
  {
    id: "word-rewrite",
    title: "Rewrite Text",
    description: "Rewrite selected text professionally",
    prompt: "Rewrite the following text to be more professional, clear, and concise. Maintain the original meaning but improve the language, structure, and flow. Provide the rewritten version.",
    category: "writing",
    hostApp: "Word",
    icon: "✏️",
  },
  {
    id: "word-summarize",
    title: "Summarize Document",
    description: "Create a concise summary",
    prompt: "Summarize this text concisely. Provide: a one-paragraph executive summary, followed by key points as bullet points. Keep it under 200 words.",
    category: "writing",
    hostApp: "Word",
    icon: "📝",
  },
  {
    id: "word-grammar",
    title: "Fix Grammar & Style",
    description: "Correct grammar, spelling, and style",
    prompt: "Review this text for grammar, spelling, punctuation, and style errors. For each issue: quote the problematic text, explain the error, and provide the correction. Then provide the fully corrected version.",
    category: "writing",
    hostApp: "Word",
    icon: "✅",
  },
  {
    id: "word-expand",
    title: "Expand Content",
    description: "Add more detail and depth",
    prompt: "Expand this text by adding more detail, examples, and depth while maintaining the same tone and style. Aim for roughly 2-3x the original length.",
    category: "writing",
    hostApp: "Word",
    icon: "📖",
  },
  {
    id: "word-shorten",
    title: "Shorten Content",
    description: "Make text more concise",
    prompt: "Condense this text to about half its length while preserving all key information and meaning. Remove redundancy, simplify sentences, and tighten the prose.",
    category: "writing",
    hostApp: "Word",
    icon: "✂️",
  },
  {
    id: "word-formal",
    title: "Make Formal",
    description: "Convert to formal/business tone",
    prompt: "Rewrite this text in a formal, professional business tone suitable for official correspondence, reports, or executive communication. Maintain accuracy.",
    category: "writing",
    hostApp: "Word",
    icon: "👔",
  },
  {
    id: "word-casual",
    title: "Make Casual",
    description: "Convert to friendly, casual tone",
    prompt: "Rewrite this text in a friendly, conversational tone suitable for blog posts, newsletters, or informal communication. Keep it engaging and approachable.",
    category: "writing",
    hostApp: "Word",
    icon: "😊",
  },
  {
    id: "word-report",
    title: "Generate Report",
    description: "Create a structured report",
    prompt: "Generate a professional report based on this content. Include: title, executive summary, introduction, main sections with headings, findings/analysis, conclusions, and recommendations. Format with proper headings.",
    category: "writing",
    hostApp: "Word",
    icon: "📄",
  },
  {
    id: "word-email",
    title: "Draft Email",
    description: "Create a professional email",
    prompt: "Draft a professional email based on this context. Include: subject line, greeting, clear body with purpose, action items, and a professional closing. Keep it concise.",
    category: "writing",
    hostApp: "Word",
    icon: "📧",
  },
  {
    id: "word-meeting-notes",
    title: "Meeting Notes",
    description: "Format as meeting minutes",
    prompt: "Format this content as professional meeting minutes. Include: date, attendees (if mentioned), agenda items, discussion points, decisions made, action items with owners and deadlines, and next steps.",
    category: "writing",
    hostApp: "Word",
    icon: "📋",
  },

  // ---- PowerPoint: Presentation ----
  {
    id: "ppt-improve-slide",
    title: "Improve Slide Text",
    description: "Make slide text more impactful",
    prompt: "Improve this slide text to be more impactful and presentation-ready. Use: concise bullet points, active voice, powerful verbs, and clear messaging. Keep each bullet under 10 words.",
    category: "presentation",
    hostApp: "PowerPoint",
    icon: "🎯",
  },
  {
    id: "ppt-speaker-notes",
    title: "Generate Speaker Notes",
    description: "Create detailed speaker notes",
    prompt: "Generate detailed speaker notes for this slide content. Include: what to say, key talking points, transitions to the next slide, and audience engagement tips. Write in a natural speaking style.",
    category: "presentation",
    hostApp: "PowerPoint",
    icon: "🎤",
  },
  {
    id: "ppt-script",
    title: "Presentation Script",
    description: "Write a full presentation script",
    prompt: "Write a complete presentation script for these slides. Include: opening hook, transitions between slides, key messages to emphasize, storytelling elements, and a strong closing. Aim for 2-3 minutes per slide.",
    category: "presentation",
    hostApp: "PowerPoint",
    icon: "📜",
  },
  {
    id: "ppt-generate-outline",
    title: "Generate Slide Outline",
    description: "Create a presentation structure",
    prompt: "Create a professional presentation outline from this content. Include: title slide, agenda, 5-8 content slides with bullet points, key visuals to include, and a closing slide. Format as slide-by-slide with titles and bullets.",
    category: "presentation",
    hostApp: "PowerPoint",
    icon: "📑",
  },
  {
    id: "ppt-board-ready",
    title: "Board Meeting Ready",
    description: "Adapt for executive audience",
    prompt: "Adapt this slide content for a board/executive meeting. Focus on: strategic impact, KPIs, financial implications, risks, and recommendations. Use concise, data-driven language. Remove operational details.",
    category: "presentation",
    hostApp: "PowerPoint",
    icon: "🏛️",
  },
  {
    id: "ppt-qa-prep",
    title: "Q&A Preparation",
    description: "Prepare for audience questions",
    prompt: "Based on this presentation content, prepare for Q&A. Generate: 10 likely questions the audience might ask, with recommended answers for each. Include tough/challenging questions too.",
    category: "presentation",
    hostApp: "PowerPoint",
    icon: "❓",
  },

  // ---- General / All Apps ----
  {
    id: "general-translate",
    title: "Translate",
    description: "Translate text to another language",
    prompt: "Translate the following text to [target language]. Maintain the original formatting, tone, and meaning. If there are cultural nuances, note them.",
    category: "translate",
    hostApp: "All",
    icon: "🌐",
  },
  {
    id: "general-sentiment",
    title: "Sentiment Analysis",
    description: "Analyze the sentiment of text",
    prompt: "Analyze the sentiment of this text. Provide: overall sentiment (positive/negative/neutral), confidence score, key phrases driving the sentiment, and a breakdown by paragraph/section if applicable.",
    category: "analysis",
    hostApp: "All",
    icon: "💬",
  },
  {
    id: "general-classify",
    title: "Classify & Tag",
    description: "Categorize and tag content",
    prompt: "Classify and tag this content. Suggest: primary category, subcategories, relevant tags/keywords, and a confidence level for each classification. Base categories on the content type.",
    category: "general",
    hostApp: "All",
    icon: "🏷️",
  },
  {
    id: "general-extract-entities",
    title: "Extract Entities",
    description: "Find names, dates, amounts, etc.",
    prompt: "Extract all named entities from this text. Categorize them: people, organizations, locations, dates, monetary amounts, percentages, products, and other. Format as a structured list.",
    category: "general",
    hostApp: "All",
    icon: "🔎",
  },
  {
    id: "general-explain",
    title: "Explain Simply",
    description: "Explain complex content simply",
    prompt: "Explain this content in simple, easy-to-understand language. Assume the reader has no prior knowledge of the subject. Use analogies and examples where helpful.",
    category: "general",
    hostApp: "All",
    icon: "💡",
  },
  {
    id: "general-python",
    title: "Generate Python Code",
    description: "Create Python code for data tasks",
    prompt: "Generate Python code (using pandas) that works with this data structure. Include: data loading, the requested transformation/analysis, and output formatting. Add comments explaining each step.",
    category: "code",
    hostApp: "All",
    icon: "🐍",
  },

  // ---- Cross-App: CA Study Assistant & Accounting ----
  {
    id: "ca-study-assistant",
    title: "CA Study Assistant",
    description: "Get help with CA exam topics and concepts",
    prompt: "You are a Chartered Accountancy (CA) study assistant. Help me understand the following CA exam topic. Provide: a clear concept explanation, key points to remember, common exam question patterns, relevant sections/standards, and mnemonic tricks if applicable. Topic: [describe your CA study topic or question]",
    category: "accounting",
    hostApp: "All",
    icon: "🎓",
  },
  {
    id: "ca-practice-question",
    title: "CA Practice Question",
    description: "Generate and solve CA practice questions",
    prompt: "Generate a CA exam-style practice question on the following topic, then provide a detailed step-by-step solution. Include: the question, relevant standard/section references, a worked solution with proper workings, and common mistakes to avoid. Topic: [describe the CA subject or chapter]",
    category: "accounting",
    hostApp: "All",
    icon: "📝",
  },
  {
    id: "ifrs-term-explainer",
    title: "IAS/IFRS Term Explainer",
    description: "Explain IAS/IFRS terms and standards",
    prompt: "Explain the following IAS/IFRS accounting term or standard in simple, clear language. Include: the full standard name and number, its purpose and scope, key definitions, recognition and measurement criteria, disclosure requirements, practical examples, and any recent amendments. Term/Standard: [enter the IAS/IFRS term or standard number]",
    category: "accounting",
    hostApp: "All",
    icon: "📚",
  },
  {
    id: "ifrs-journal-entry",
    title: "IFRS Journal Entry Helper",
    description: "Generate journal entries under IAS/IFRS",
    prompt: "Help me prepare the correct journal entries under IAS/IFRS for the following transaction. Include: the journal entry with proper debits and credits, the relevant IAS/IFRS standard reference, explanation of the accounting treatment, subsequent measurement entries if applicable, and disclosure notes. Transaction: [describe the accounting transaction]",
    category: "accounting",
    hostApp: "All",
    icon: "📋",
  },
  {
    id: "ifrs-vs-gaap",
    title: "IFRS vs GAAP Comparison",
    description: "Compare IAS/IFRS and US GAAP treatments",
    prompt: "Compare the IAS/IFRS and US GAAP treatment for the following accounting topic. Include: key similarities, key differences, relevant standard numbers from both frameworks, practical examples showing how the same transaction would be recorded differently, and which treatment is more conservative. Topic: [describe the accounting area]",
    category: "accounting",
    hostApp: "All",
    icon: "⚖️",
  },
];

/** Get templates filtered by host app */
export function getTemplatesForHost(host: OfficeHostType): PromptTemplate[] {
  return PROMPT_TEMPLATES.filter(
    (t) => t.hostApp === host || t.hostApp === "All"
  );
}

/** Get templates by category */
export function getTemplatesByCategory(
  category: string,
  host?: OfficeHostType
): PromptTemplate[] {
  let templates = PROMPT_TEMPLATES.filter((t) => t.category === category);
  if (host) {
    templates = templates.filter((t) => t.hostApp === host || t.hostApp === "All");
  }
  return templates;
}

/** Search templates */
export function searchTemplates(query: string, host?: OfficeHostType): PromptTemplate[] {
  const q = query.toLowerCase();
  let templates = PROMPT_TEMPLATES.filter(
    (t) =>
      t.title.toLowerCase().includes(q) ||
      t.description.toLowerCase().includes(q) ||
      t.prompt.toLowerCase().includes(q)
  );
  if (host) {
    templates = templates.filter((t) => t.hostApp === host || t.hostApp === "All");
  }
  return templates;
}
