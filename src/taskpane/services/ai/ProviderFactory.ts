import { BaseAIProvider } from './types';
import { NvidiaProvider } from './NvidiaProvider';
import { OpenAIProvider } from './OpenAIProvider';
import { AnthropicProvider } from './AnthropicProvider';
import { GeminiProvider } from './GeminiProvider';
import { GroqProvider } from './GroqProvider';
import { OpenRouterProvider } from './OpenRouterProvider';
import { OllamaProvider } from './OllamaProvider';
import { AIProviderType } from '../../types'; // Ensure this matches actual location or just use string if type doesn't exist yet

export const PROVIDER_CONFIGS: Record<string, {name: string, color: string, icon: string, description: string}> = {
  nvidia: { name: 'NVIDIA NIM', color: '#76B900', icon: 'Bot24Regular', description: 'NVIDIA API' },
  openai: { name: 'OpenAI', color: '#10A37F', icon: 'Bot24Regular', description: 'OpenAI API' },
  anthropic: { name: 'Anthropic', color: '#D4A574', icon: 'Bot24Regular', description: 'Claude API' },
  gemini: { name: 'Google Gemini', color: '#4285F4', icon: 'Bot24Regular', description: 'Gemini API' },
  groq: { name: 'Groq', color: '#F55036', icon: 'Bot24Regular', description: 'Fast Llama Inference' },
  openrouter: { name: 'OpenRouter', color: '#6366F1', icon: 'Bot24Regular', description: 'OpenRouter API' },
  ollama: { name: 'Ollama (Local)', color: '#FFFFFF', icon: 'Bot24Regular', description: 'Local Models' }
};

const providers: Map<string, BaseAIProvider> = new Map();

// Initialize standard providers
providers.set('nvidia', new NvidiaProvider());
providers.set('openai', new OpenAIProvider());
providers.set('anthropic', new AnthropicProvider());
providers.set('gemini', new GeminiProvider());
providers.set('groq', new GroqProvider());
providers.set('openrouter', new OpenRouterProvider());
providers.set('ollama', new OllamaProvider());

export function getProvider(type: string | AIProviderType): BaseAIProvider {
  const provider = providers.get(type);
  if (!provider) {
    throw new Error(`Provider ${type} not found`);
  }
  return provider;
}

export function getAllProviders(): BaseAIProvider[] {
  return Array.from(providers.values());
}

export function getProviderConfig(type: string | AIProviderType) {
  const provider = getProvider(type);
  const config = PROVIDER_CONFIGS[type];
  return {
    id: provider.id,
    name: provider.name,
    color: config?.color || '#000000',
    icon: config?.icon || 'Bot24Regular',
    requiresKey: provider.requiresKey()
  };
}
