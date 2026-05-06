import type { LanguageModel } from 'ai';

type GeminiCliOpts = { authType?: 'oauth-personal' | 'gemini-api-key'; apiKey?: string };

export function makeGeminiCli(opts: GeminiCliOpts): (modelId: string) => LanguageModel {
  const authType = opts.authType ?? 'oauth-personal';
  return (modelId) => {
    let _inner: LanguageModel | null = null;

    const load = async (): Promise<LanguageModel> => {
      if (_inner) return _inner;
      const mod = await import('ai-sdk-provider-gemini-cli').catch((err: Error) => {
        throw new Error(`Gemini CLI provider unavailable (is the gemini CLI installed?): ${err.message}`);
      });
      const provider =
        authType === 'gemini-api-key'
          ? mod.createGeminiProvider({ authType: 'gemini-api-key', ...(opts.apiKey ? { apiKey: opts.apiKey } : {}) })
          : mod.createGeminiProvider({ authType: 'oauth-personal' });
      _inner = provider(modelId);
      return _inner;
    };

    return {
      specificationVersion: 'v2' as const,
      provider: 'gemini-cli',
      modelId,
      defaultObjectGenerationMode: undefined,
      doGenerate: async (options: never) => (await load()).doGenerate(options),
      doStream: async (options: never) => (await load()).doStream(options),
    } as unknown as LanguageModel;
  };
}
