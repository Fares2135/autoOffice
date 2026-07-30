import { describe, it, expect } from 'vitest';
import {
  thinkingConfigFor,
  thinkingProviderOptions,
  supportsThinkingControl,
  availableLevels,
  nearestLevel,
} from './thinking.ts';

describe('thinkingConfigFor — Gemini 3.x uses named levels', () => {
  it('passes a supported level straight through', () => {
    expect(thinkingConfigFor('gemini-3-flash', 'medium')).toEqual({ thinkingLevel: 'medium' });
    expect(thinkingConfigFor('gemini-3-flash', 'minimal')).toEqual({ thinkingLevel: 'minimal' });
  });

  it('clamps to the nearest level a model actually accepts', () => {
    // Gemini 3 Pro only has low | high.
    expect(thinkingConfigFor('gemini-3-pro', 'minimal')).toEqual({ thinkingLevel: 'low' });
    expect(thinkingConfigFor('gemini-3-pro', 'medium')).toEqual({ thinkingLevel: 'low' });
    expect(thinkingConfigFor('gemini-3-pro', 'high')).toEqual({ thinkingLevel: 'high' });
  });

  it('knows 3.1 Pro added medium but still has no minimal', () => {
    expect(thinkingConfigFor('gemini-3.1-pro', 'medium')).toEqual({ thinkingLevel: 'medium' });
    expect(thinkingConfigFor('gemini-3.1-pro', 'minimal')).toEqual({ thinkingLevel: 'low' });
  });

  it('never emits a numeric budget for a Gemini 3 model', () => {
    for (const level of ['minimal', 'low', 'medium', 'high'] as const) {
      expect(thinkingConfigFor('gemini-3-pro', level)!.thinkingBudget).toBeUndefined();
    }
  });
});

describe('thinkingConfigFor — Gemini 2.5 uses a token budget', () => {
  it('maps the four levels onto budgets', () => {
    expect(thinkingConfigFor('gemini-2.5-flash', 'low')).toEqual({ thinkingBudget: 1024 });
    expect(thinkingConfigFor('gemini-2.5-flash', 'medium')).toEqual({ thinkingBudget: 8192 });
    expect(thinkingConfigFor('gemini-2.5-flash', 'high')).toEqual({ thinkingBudget: 24576 });
  });

  it('switches thinking off on Flash, which allows zero', () => {
    expect(thinkingConfigFor('gemini-2.5-flash', 'minimal')).toEqual({ thinkingBudget: 0 });
  });

  it('falls back to the floor on Pro, which cannot disable thinking', () => {
    expect(thinkingConfigFor('gemini-2.5-pro', 'minimal')).toEqual({ thinkingBudget: 128 });
  });

  it('respects the higher Pro ceiling', () => {
    expect(thinkingConfigFor('gemini-2.5-pro', 'high')).toEqual({ thinkingBudget: 24576 });
  });

  it('keeps flash-lite above its 512 floor for any thinking level', () => {
    expect(thinkingConfigFor('gemini-2.5-flash-lite', 'low')).toEqual({ thinkingBudget: 1024 });
    expect(thinkingConfigFor('gemini-2.5-flash-lite', 'minimal')).toEqual({ thinkingBudget: 0 });
  });

  it('never emits a named level for a 2.5 model', () => {
    expect(thinkingConfigFor('gemini-2.5-flash', 'high')!.thinkingLevel).toBeUndefined();
  });
});

describe('thinkingConfigFor — nothing to send', () => {
  it('sends nothing on auto', () => {
    expect(thinkingConfigFor('gemini-3-pro', 'auto')).toBeNull();
    expect(thinkingConfigFor('gemini-2.5-flash', 'auto')).toBeNull();
  });

  it('sends nothing for a model with no thinking knob', () => {
    expect(thinkingConfigFor('gemini-1.5-pro', 'high')).toBeNull();
    expect(thinkingConfigFor('', 'high')).toBeNull();
  });
});

describe('thinkingProviderOptions', () => {
  it('wraps the config under the google provider key', () => {
    expect(thinkingProviderOptions('google', 'gemini-3-flash', 'high'))
      .toEqual({ google: { thinkingConfig: { thinkingLevel: 'high' } } });
  });

  it('is undefined for every non-Google provider', () => {
    expect(thinkingProviderOptions('anthropic', 'claude-opus-5', 'high')).toBeUndefined();
    expect(thinkingProviderOptions('openrouter', 'google/gemini-3-pro', 'high')).toBeUndefined();
  });

  it('is undefined on auto, so today\'s behaviour is unchanged by default', () => {
    expect(thinkingProviderOptions('google', 'gemini-3-flash', 'auto')).toBeUndefined();
  });
});

describe('supportsThinkingControl', () => {
  it('is true only for Gemini models on the Google provider', () => {
    expect(supportsThinkingControl('google', 'gemini-3-pro')).toBe(true);
    expect(supportsThinkingControl('google', 'gemini-2.5-flash')).toBe(true);
    expect(supportsThinkingControl('google', 'gemini-1.5-pro')).toBe(false);
    expect(supportsThinkingControl('anthropic', 'gemini-3-pro')).toBe(false);
    expect(supportsThinkingControl('google', undefined)).toBe(false);
  });
});

describe('availableLevels', () => {
  it('hides levels a model rejects', () => {
    expect(availableLevels('gemini-3-pro')).toEqual(['auto', 'low', 'high']);
    expect(availableLevels('gemini-3.1-pro')).toEqual(['auto', 'low', 'medium', 'high']);
    expect(availableLevels('gemini-3-flash')).toEqual(['auto', 'minimal', 'low', 'medium', 'high']);
  });

  it('offers everything for budget-based models', () => {
    expect(availableLevels('gemini-2.5-flash')).toEqual(['auto', 'minimal', 'low', 'medium', 'high']);
  });
});

describe('nearestLevel', () => {
  it('breaks ties downward', () => {
    expect(nearestLevel('medium', ['low', 'high'])).toBe('low');
  });
});
