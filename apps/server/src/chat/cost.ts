import type { LanguageModelUsage } from 'ai';
import { isLocalProvider, type ProviderKind, type UsageCost } from '@autooffice/shared';

type Rate = { input: number; output: number };

// USD per million tokens. This is an intentionally small snapshot inherited
// from the previous AutoOffice cost tracker; unknown models stay tokens-only.
export const PRICING_VERSION = '2026-05';
const RATES: Record<string, Rate> = {
  'claude-opus-4-7': { input: 5, output: 25 },
  'claude-sonnet-4-6': { input: 3, output: 15 },
  'claude-haiku-4-5': { input: 1, output: 5 },
  'gpt-4o': { input: 2.5, output: 10 },
  'gpt-4o-mini': { input: 0.15, output: 0.6 },
  'gemini-2.5-pro': { input: 1.25, output: 10 },
  'gemini-2.5-flash': { input: 0.3, output: 2.5 },
  'llama-3.3-70b-versatile': { input: 0.59, output: 0.79 },
  'llama-3.1-8b-instant': { input: 0.05, output: 0.08 },
  'grok-3': { input: 3, output: 15 },
  'grok-3-mini': { input: 0.3, output: 0.5 },
  'deepseek-chat': { input: 0.27, output: 1.1 },
  'deepseek-reasoner': { input: 0.55, output: 2.19 },
};

export function computeUsageCost(
  providerKind: ProviderKind | null,
  modelId: string,
  usage: LanguageModelUsage,
): UsageCost {
  const inputTokens = usage.inputTokens ?? 0;
  const outputTokens = usage.outputTokens ?? 0;
  const totalTokens = usage.totalTokens ?? inputTokens + outputTokens;
  if (providerKind && isLocalProvider(providerKind)) {
    return { inputTokens, outputTokens, totalTokens, totalUsd: 0, source: 'local-free' };
  }
  const rate = RATES[modelId];
  if (!rate) {
    return { inputTokens, outputTokens, totalTokens, totalUsd: null, source: 'tokens-only' };
  }
  return {
    inputTokens,
    outputTokens,
    totalTokens,
    totalUsd: (inputTokens * rate.input + outputTokens * rate.output) / 1_000_000,
    source: 'estimated',
  };
}
