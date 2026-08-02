// ============================================================
// Office AI Copilot — Core Type Definitions
// ============================================================

// ---- AI Provider Types ----

export type AIProviderType =
  | "nvidia"
  | "openai"
  | "anthropic"
  | "gemini"
  | "groq"
  | "openrouter"
  | "ollama";

export interface AIProviderConfig {
  id: AIProviderType;
  name: string;
  description: string;
  baseUrl: string;
  requiresApiKey: boolean;
  defaultModel: string;
  models: AIModel[];
  icon: string;
  color: string;
}

export interface AIModel {
  id: string;
  name: string;
  provider: AIProviderType;
  description?: string;
  contextWindow?: number;
  maxTokens?: number;
}

/** A user-defined model entry with its own key and provider */
export interface CustomModel {
  id: string;           // unique UUID
  name: string;         // user-friendly display name
  modelId: string;      // actual API model ID (e.g. "meta/llama-3.1-70b-instruct")
  provider: AIProviderType;
  apiKey: string;       // API key stored per model
  baseUrl?: string;     // optional override (for Ollama/OpenRouter)
}

export interface ChatMessage {
  id: string;
  role: "user" | "assistant" | "system" | "compaction_summary";
  content: string;
  timestamp: number;
  model?: string;
  provider?: AIProviderType;
  /** If the message included cell/document context */
  contextIncluded?: boolean;
  /** Tokens used for this message */
  tokens?: number;
  /** If the AI returned a reasoning/thinking trace (Claude, o1, DeepSeek-R1) */
  thinking?: string;
  /** Whether this message has been hidden from the LLM context by compaction */
  compacted?: boolean;
}

export interface ChatConversation {
  id: string;
  title: string;
  messages: ChatMessage[];
  createdAt: number;
  updatedAt: number;
  hostApp: OfficeHostType;
  model?: string;
  provider?: AIProviderType;
}

export interface StreamingState {
  isStreaming: boolean;
  currentText: string;
  abortController: AbortController | null;
}

export interface ChatOptions {
  model: string;
  temperature?: number;
  maxTokens?: number;
  systemPrompt?: string;
  stream?: boolean;
}

// ---- Office Types ----

export type OfficeHostType = "Excel" | "Word" | "PowerPoint" | "Unknown" | "All";

export interface OfficeContextInfo {
  host: OfficeHostType;
  platform: string;
  version: string;
  isReady: boolean;
  /** Which API sets are supported */
  supportedApis: SupportedApiInfo;
}

export interface SupportedApiInfo {
  // Excel API versions
  excelApi1_1: boolean;
  excelApi1_2: boolean;
  excelApi1_3: boolean;
  excelApi1_4: boolean;
  excelApi1_5: boolean;
  excelApi1_6: boolean;
  excelApi1_7: boolean;
  excelApi1_8: boolean;
  excelApi1_9: boolean;
  excelApi1_10: boolean;
  excelApi1_11: boolean;
  excelApi1_12: boolean;
  excelApi1_13: boolean;
  excelApi1_14: boolean;
  // Word API versions
  wordApi1_1: boolean;
  wordApi1_2: boolean;
  wordApi1_3: boolean;
  wordApi1_4: boolean;
  // PowerPoint API versions
  powerPointApi1_1: boolean;
  powerPointApi1_2: boolean;
  powerPointApi1_3: boolean;
  powerPointApi1_4: boolean;
  powerPointApi1_5: boolean;
}

/** Data extracted from Excel selection or range */
export interface CellData {
  address: string;
  values: (string | number | boolean)[][];
  formulas?: string[][];
  numberFormat?: string[][];
  rowCount: number;
  columnCount: number;
  sheetName: string;
}

/** Data extracted from Word document */
export interface DocumentData {
  selectedText: string;
  fullText?: string;
  paragraphCount?: number;
  wordCount?: number;
  title?: string;
}

/** Data extracted from PowerPoint slides */
export interface SlideData {
  slideIndex: number;
  slideCount: number;
  currentSlideText: string;
  allSlidesText?: string[];
  title?: string;
  notes?: string;
}

/** Workbook metadata for context */
export interface WorkbookInfo {
  name: string;
  sheets: SheetInfo[];
  activeSheet: string;
  namedRanges?: string[];
}

export interface SheetInfo {
  name: string;
  rowCount?: number;
  columnCount?: number;
  usedRangeAddress?: string;
  isVisible: boolean;
}

// ---- App State Types ----

export type AppView = "chat" | "settings";

export type ThemeMode = "light" | "dark" | "auto";

export interface AppSettings {
  // AI Configuration
  activeProvider: AIProviderType;
  activeModel: string;
  activeCustomModelId: string;   // UUID of the active CustomModel
  customModels: CustomModel[];   // user-defined models
  apiKeys: Partial<Record<AIProviderType, string>>;

  // UI Preferences
  theme: ThemeMode;
  fontSize: "small" | "medium" | "large";
  compactMode: boolean;

  // Privacy
  includeContextByDefault: boolean;
  maxContextCells: number;
  maxContextCharacters: number;

  // Chat
  streamResponses: boolean;
  saveConversations: boolean;
  maxConversationHistory: number;
}

export interface AppState {
  // Navigation
  currentView: AppView;
  officeContext: OfficeContextInfo;

  // Chat
  conversations: ChatConversation[];
  activeConversationId: string | null;
  streaming: StreamingState;

  // Settings
  settings: AppSettings;

  // UI State
  isLoading: boolean;
  error: string | null;
  notification: AppNotification | null;
}

export interface AppNotification {
  id: string;
  type: "success" | "error" | "warning" | "info";
  message: string;
  duration?: number;
}

// ---- Export Types ----

export type ExportFormat = "markdown" | "json" | "text" | "html";

// ---- Prompt Types ----

export type PromptCategory =
  | "analysis"
  | "cleaning"
  | "formula"
  | "formatting"
  | "automation"
  | "general"
  | "audit"
  | "presentation"
  | "reporting"
  | "ca"
  | "compliance-bd"
  | "writing"
  | "translate"
  | "chart"
  | "finance"
  | "code"
  | "accounting"
  | "all";

export interface PromptTemplate {
  id: string;
  title: string;
  description: string;
  prompt: string;
  category: PromptCategory;
  hostApp: OfficeHostType;
  icon: string;
}

