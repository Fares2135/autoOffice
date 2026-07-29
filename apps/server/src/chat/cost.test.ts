import { describe, expect, it } from 'vitest';
import { computeUsageCost } from './cost';

const usage = { inputTokens: 1_000, outputTokens: 500, totalTokens: 1_500 };

describe('computeUsageCost', () => {
  it('marks local providers as free while retaining usage', () => {
    expect(computeUsageCost('lmstudio', 'local', usage as any)).toEqual({
      inputTokens: 1_000,
      outputTokens: 500,
      totalTokens: 1_500,
      totalUsd: 0,
      source: 'local-free',
    });
  });

  it('estimates known models and leaves unknown prices unset', () => {
    expect(computeUsageCost('openai', 'gpt-4o-mini', usage as any)).toMatchObject({
      totalUsd: 0.00045,
      source: 'estimated',
    });
    expect(computeUsageCost('openai', 'future-model', usage as any)).toMatchObject({
      totalUsd: null,
      source: 'tokens-only',
    });
  });
});
