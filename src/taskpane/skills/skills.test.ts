import { describe, it, expect } from 'vitest';
import { summarize, describeSkills, lookupSkills, listSkills } from './index.ts';

describe('summarize', () => {
  it('takes the first prose sentence, not the heading', () => {
    const md = '# Tables\n\nCreate and edit tables. More detail follows.\n';
    expect(summarize(md)).toBe('Create and edit tables.');
  });

  it('rejoins a sentence wrapped across markdown lines', () => {
    const md = '# X\n\nUse this skill for Arabic or Hebrew text, mixed\ncontent, and tables. Next sentence.\n';
    expect(summarize(md)).toBe('Use this skill for Arabic or Hebrew text, mixed content, and tables.');
  });

  it('skips fenced code and returns empty for a doc with no prose', () => {
    expect(summarize('# X\n\n```js\ncode();\n```\n')).toBe('');
    expect(summarize('')).toBe('');
  });

  it('truncates an over-long sentence', () => {
    const md = `# X\n\n${'word '.repeat(60)}.\n`;
    const s = summarize(md);
    expect(s.length).toBeLessThanOrEqual(111);
    expect(s.endsWith('…')).toBe(true);
  });
});

describe('describeSkills', () => {
  it('emits one line per registered skill for every host', () => {
    for (const host of ['word', 'excel', 'powerpoint'] as const) {
      const lines = describeSkills(host).split('\n');
      expect(lines).toHaveLength(listSkills(host).length);
      for (const line of lines) expect(line.startsWith('- ')).toBe(true);
    }
  });

  it('gives the real Word skills a summary, not a bare name', () => {
    const text = describeSkills('word');
    expect(text).toMatch(/- tables — \S/);
    expect(text).toMatch(/- arabic-rtl — \S/);
  });
});

describe('lookupSkills', () => {
  it('concatenates several domains under labelled headings', () => {
    const out = lookupSkills('word', ['tables', 'styles']);
    expect(out).toContain('## Skill: tables');
    expect(out).toContain('## Skill: styles');
  });

  it('reports unknown names without throwing', () => {
    expect(lookupSkills('word', ['nope'])).toContain('not found');
  });

  it('handles an empty list', () => {
    expect(lookupSkills('word', [])).toContain('not found');
  });
});
