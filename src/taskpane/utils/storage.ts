// ============================================================
// Local Storage Wrapper — Persists settings, keys, conversations
// ============================================================

import { AppSettings, ChatConversation, AIProviderType } from "../types";

const STORAGE_PREFIX = "office-ai-copilot";
const PROVIDERS: AIProviderType[] = [
  "nvidia",
  "openai",
  "anthropic",
  "gemini",
  "groq",
  "openrouter",
  "ollama",
];

const KEYS = {
  settings: `${STORAGE_PREFIX}:settings`,
  conversations: `${STORAGE_PREFIX}:conversations`,
  favorites: `${STORAGE_PREFIX}:favorites`,
  apiKey: (provider: AIProviderType) => `${STORAGE_PREFIX}:apikey:${provider}`,
} as const;

// ---- Generic Helpers ----

function getItem<T>(key: string, fallback: T): T {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return fallback;
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

function setItem<T>(key: string, value: T): boolean {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch (e) {
    console.warn("[Storage] Failed to save:", key, e);
    // Surface the error so the UI can notify the user — otherwise their
    // data appears saved (React state is updated) but isn't actually
    // persisted, and they only find out on refresh.
    if (typeof window !== 'undefined') {
      window.dispatchEvent(new CustomEvent('storage-error', {
        detail: { key, error: (e as Error).message },
      }));
    }
    return false;
  }
}

function removeItem(key: string): void {
  try {
    localStorage.removeItem(key);
  } catch (e) {
    console.warn("[Storage] Failed to remove:", key, e);
  }
}

// ---- Settings ----

export const DEFAULT_SETTINGS: AppSettings = {
  activeProvider: "nvidia",
  activeModel: "meta/llama-3.1-70b-instruct",
  activeCustomModelId: "",
  customModels: [],
  apiKeys: {},
  theme: "auto",
  fontSize: "medium",
  compactMode: false,
  includeContextByDefault: true,
  maxContextCells: 500,
  maxContextCharacters: 8000,
  streamResponses: true,
  saveConversations: true,
  maxConversationHistory: 50,
};

export function loadSettings(): AppSettings {
  const saved = getItem<Partial<AppSettings>>(KEYS.settings, {});
  // Merge with defaults so new settings are always present
  const settings = { ...DEFAULT_SETTINGS, ...saved };

  // Load API keys separately (stored individually for safety)
  const apiKeys: Partial<Record<AIProviderType, string>> = {};
  for (const p of PROVIDERS) {
    const key = getItem<string>(KEYS.apiKey(p), "");
    if (key) apiKeys[p] = key;
  }
  settings.apiKeys = apiKeys;

  return settings;
}

export function saveSettings(settings: AppSettings): void {
  // Save API keys separately
  const { apiKeys, ...rest } = settings;
  setItem(KEYS.settings, rest);

  if (apiKeys) {
    for (const [provider, key] of Object.entries(apiKeys)) {
      if (key) {
        setItem(KEYS.apiKey(provider as AIProviderType), key);
      } else {
        removeItem(KEYS.apiKey(provider as AIProviderType));
      }
    }
  }
  // Ensure removed/reset keys are deleted from storage even if apiKeys is empty.
  for (const provider of PROVIDERS) {
    if (!apiKeys?.[provider]) {
      removeItem(KEYS.apiKey(provider));
    }
  }
}

// ---- Conversations ----

export function loadConversations(): ChatConversation[] {
  return getItem<ChatConversation[]>(KEYS.conversations, []);
}

export function saveConversations(conversations: ChatConversation[]): void {
  setItem(KEYS.conversations, conversations);
}

export function deleteConversation(id: string): ChatConversation[] {
  const convos = loadConversations().filter((c) => c.id !== id);
  saveConversations(convos);
  return convos;
}

export function clearAllConversations(): void {
  saveConversations([]);
}

// ---- Favorites ----

export function loadFavoritePrompts(): string[] {
  return getItem<string[]>(KEYS.favorites, []);
}

export function saveFavoritePrompts(ids: string[]): void {
  setItem(KEYS.favorites, ids);
}

export function toggleFavorite(promptId: string): string[] {
  const favs = loadFavoritePrompts();
  const idx = favs.indexOf(promptId);
  if (idx >= 0) {
    favs.splice(idx, 1);
  } else {
    favs.push(promptId);
  }
  saveFavoritePrompts(favs);
  return favs;
}

// ---- Utility ----

export function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;
}

export function clearAllData(): void {
  const keysToRemove: string[] = [];
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key?.startsWith(STORAGE_PREFIX)) {
      keysToRemove.push(key);
    }
  }
  keysToRemove.forEach((k) => localStorage.removeItem(k));
}
