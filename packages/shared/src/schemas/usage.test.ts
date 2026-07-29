import { describe, expect, it } from 'vitest';
import { sumUsageCosts, UsageCostSchema } from './usage';

describe('UsageCost', () => {
  it('sums estimated usage and cost', () => {
    expect(
      sumUsageCosts([
        { inputTokens: 10, outputTokens: 5, totalTokens: 15, totalUsd: 0.01, source: 'estimated' },
        { inputTokens: 4, outputTokens: 1, totalTokens: 5, totalUsd: 0.02, source: 'estimated' },
      ]),
    ).toEqual({
      inputTokens: 14,
      outputTokens: 6,
      totalTokens: 20,
      totalUsd: 0.03,
      source: 'estimated',
    });
  });

  it('keeps unknown prices honest while preserving token totals', () => {
    const result = sumUsageCosts([
      { inputTokens: 10, outputTokens: 5, totalTokens: 15, totalUsd: null, source: 'tokens-only' },
    ]);
    expect(UsageCostSchema.parse(result)).toMatchObject({ totalTokens: 15, totalUsd: null });
  });
});
