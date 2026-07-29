import * as z from 'zod';

export const UsageCostSchema = z.object({
  inputTokens: z.number().int().nonnegative(),
  outputTokens: z.number().int().nonnegative(),
  totalTokens: z.number().int().nonnegative(),
  totalUsd: z.number().nonnegative().nullable(),
  source: z.enum(['estimated', 'tokens-only', 'local-free']),
});

export type UsageCost = z.infer<typeof UsageCostSchema>;

export function sumUsageCosts(costs: readonly UsageCost[]): UsageCost | null {
  if (costs.length === 0) return null;
  const hasUnknownPrice = costs.some((cost) => cost.source === 'tokens-only');
  const allLocal = costs.every((cost) => cost.source === 'local-free');
  return {
    inputTokens: costs.reduce((sum, cost) => sum + cost.inputTokens, 0),
    outputTokens: costs.reduce((sum, cost) => sum + cost.outputTokens, 0),
    totalTokens: costs.reduce((sum, cost) => sum + cost.totalTokens, 0),
    totalUsd: hasUnknownPrice
      ? null
      : costs.reduce((sum, cost) => sum + (cost.totalUsd ?? 0), 0),
    source: allLocal ? 'local-free' : hasUnknownPrice ? 'tokens-only' : 'estimated',
  };
}
