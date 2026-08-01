import { OpenAICompatibleProvider } from './OpenAICompatibleProvider';

export class OllamaProvider extends OpenAICompatibleProvider {
  readonly id = 'ollama';
  readonly name = 'Ollama (Local)';
  readonly baseUrl = 'http://localhost:11434/v1';

  /** Ollama runs locally and supports CORS — skip the proxy to avoid
   *  routing localhost traffic through a remote Vercel function. */
  protected useProxy = false;

  requiresKey(): boolean { return false; }

  private dynamicModels: Array<{ id: string; name: string }> = [];

  constructor() {
    super();
    this.fetchLocalModels().catch((e) => console.warn('Could not fetch Ollama models, using fallback list.', e));
  }

  private async fetchLocalModels() {
    const res = await fetch('http://localhost:11434/api/tags');
    if (!res.ok) return;
    const data = await res.json();
    if (Array.isArray(data.models)) {
      this.dynamicModels = data.models.map((m: { name: string }) => ({
        id: m.name,
        name: m.name,
      }));
    }
  }

  getModels() {
    if (this.dynamicModels.length > 0) return this.dynamicModels;
    return [
      { id: 'llama3.1', name: 'Llama 3.1' },
      { id: 'mistral', name: 'Mistral' },
      { id: 'phi3', name: 'Phi-3' },
      { id: 'codellama', name: 'Code Llama' },
    ];
  }
}
