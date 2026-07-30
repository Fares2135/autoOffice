// Reasoning-depth control for Gemini.
//
// Two incompatible schemes live behind one UI setting: Gemini 3.x takes a
// named thinkingLevel, Gemini 2.5 takes a numeric thinkingBudget. Not every
// model accepts every level either — Gemini 3 Pro has no 'minimal', and
// Gemini 2.5 Pro cannot switch thinking off at all — so a request is clamped
// to what the chosen model actually supports rather than rejected by the API.
//
// ponytail: Gemini only for now, on purpose. Other providers spell reasoning
// effort differently; add them when someone asks, behind the same setting.

export type ThinkingLevel = 'auto' | 'minimal' | 'low' | 'medium' | 'high';

/** Order matters: used to find the nearest level a model supports. */
export const THINKING_LEVELS: readonly ThinkingLevel[] = ['auto', 'minimal', 'low', 'medium', 'high'];

const NAMED: readonly Exclude<ThinkingLevel, 'auto'>[] = ['minimal', 'low', 'medium', 'high'];

export interface ThinkingConfig {
  thinkingLevel?: string;
  thinkingBudget?: number;
  // Index signature so the object satisfies the AI SDK's JSON-shaped
  // providerOptions without a cast at the call site.
  [key: string]: string | number | undefined;
}

/** Budget scheme limits per Gemini 2.5 family. */
interface BudgetRules {
  min: number;
  max: number;
  canDisable: boolean;
}

const BUDGET_TARGETS: Record<Exclude<ThinkingLevel, 'auto'>, number> = {
  minimal: 0,
  low: 1024,
  medium: 8192,
  high: 24576,
};

function budgetRules(model: string): BudgetRules | null {
  if (!/gemini-2\.5/i.test(model)) return null;
  if (/flash-lite/i.test(model)) return { min: 512, max: 24576, canDisable: true };
  if (/flash/i.test(model)) return { min: 1, max: 24576, canDisable: true };
  if (/pro/i.test(model)) return { min: 128, max: 32768, canDisable: false };
  return { min: 1, max: 24576, canDisable: true };
}

/** Named levels each Gemini 3.x model accepts. */
function namedLevels(model: string): readonly string[] | null {
  if (!/gemini-3/i.test(model)) return null;
  if (/flash/i.test(model)) return ['minimal', 'low', 'medium', 'high'];
  if (/3\.1.*pro|pro.*3\.1/i.test(model)) return ['low', 'medium', 'high'];
  if (/pro/i.test(model)) return ['low', 'high'];
  return ['low', 'high'];
}

/** Closest supported level to what was asked; ties resolve downward (cheaper). */
export function nearestLevel(requested: Exclude<ThinkingLevel, 'auto'>, supported: readonly string[]): string {
  if (supported.includes(requested)) return requested;
  const want = NAMED.indexOf(requested);
  let best = supported[0];
  let bestDistance = Number.POSITIVE_INFINITY;
  for (const candidate of supported) {
    const distance = Math.abs(NAMED.indexOf(candidate as typeof NAMED[number]) - want);
    if (distance < bestDistance) {
      best = candidate;
      bestDistance = distance;
    }
  }
  return best;
}

/**
 * Translates the UI setting into a thinkingConfig for the given model, or
 * null when nothing should be sent — 'auto' leaves the model at its own
 * default, and a non-Gemini model has no equivalent knob.
 */
export function thinkingConfigFor(model: string, level: ThinkingLevel): ThinkingConfig | null {
  if (level === 'auto' || !model) return null;

  const named = namedLevels(model);
  if (named) return { thinkingLevel: nearestLevel(level, named) };

  const rules = budgetRules(model);
  if (!rules) return null;

  const target = BUDGET_TARGETS[level];
  if (target === 0) {
    return { thinkingBudget: rules.canDisable ? 0 : rules.min };
  }
  return { thinkingBudget: Math.min(Math.max(target, rules.min), rules.max) };
}

/** Ready-to-spread providerOptions for streamText, or undefined for "send nothing". */
export function thinkingProviderOptions(
  providerId: string,
  model: string,
  level: ThinkingLevel,
): { google: { thinkingConfig: ThinkingConfig } } | undefined {
  if (providerId !== 'google') return undefined;
  const config = thinkingConfigFor(model, level);
  return config ? { google: { thinkingConfig: config } } : undefined;
}

/** True when the reasoning control is meaningful for the current selection. */
export function supportsThinkingControl(providerId: string | undefined, model: string | undefined): boolean {
  if (providerId !== 'google' || !model) return false;
  return namedLevels(model) !== null || budgetRules(model) !== null;
}

/** Levels worth offering for this model, always including 'auto'. */
export function availableLevels(model: string | undefined): readonly ThinkingLevel[] {
  if (!model) return THINKING_LEVELS;
  const named = namedLevels(model);
  if (!named) return THINKING_LEVELS;
  return ['auto', ...NAMED.filter((l) => named.includes(l))];
}
