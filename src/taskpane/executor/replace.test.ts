import { describe, it, expect } from 'vitest';
import { planReplacements, formatPlan, stripBidi, describeInvisibles } from './replace.ts';

const RLM = '‏';
const LRM = '‎';

describe('the exact task that failed: strip the Arabic half off index entries', () => {
  it('keeps "Module 26" and drops the rest', () => {
    const { plans } = planReplacements(
      ['Module 26 | المقطع الصوتي'],
      'Module (\\d+) \\| .*',
      'Module $1',
      { regex: true },
    );
    expect(plans).toHaveLength(1);
    expect(plans[0].after).toBe('Module 26');
    expect(plans[0].wholeParagraph).toBe(true);
  });

  it('matches through an invisible RLM, which is why the old regex never fired', () => {
    const text = `Module 26 ${RLM}| المقطع الصوتي`;
    // A plain \s*\| regex fails on this text — that is the reported bug.
    expect(/Module \d+\s*\|/.test(text)).toBe(false);

    const { plans } = planReplacements([text], 'Module (\\d+) *\\| .*', 'Module $1', { regex: true });
    expect(plans).toHaveLength(1);
    expect(plans[0].newText).toBe('Module 26');
  });

  it('CRITICAL: a paragraph holding many tab-separated entries keeps its siblings', () => {
    // This is the real index shape: one paragraph, many entries.
    const paragraph =
      '040\tModule 88 | المقطع الصوتي\t\t\t041\tModule 89 | المقطع الصوتي\t\t\t042\tModule 90 | المقطع الصوتي';
    const { plans } = planReplacements([paragraph], 'Module (\\d+) \\| [^\\t]*', 'Module $1', { regex: true });

    expect(plans).toHaveLength(3);
    // None of them is a whole-paragraph replace, so none can wipe the others.
    expect(plans.every((p) => p.wholeParagraph === false)).toBe(true);
    expect(plans[0].matchesInParagraph).toBe(3);

    // Applying the first one leaves entries 89 and 90 intact.
    expect(plans[0].after).toContain('Module 89 | المقطع الصوتي');
    expect(plans[0].after).toContain('Module 90 | المقطع الصوتي');
    expect(plans[0].after).toContain('040\tModule 88\t');
  });

  it('reports each match separately so a partial apply is reviewable', () => {
    const { plans } = planReplacements(
      ['Module 1 | صوت', 'unrelated', 'Module 2 | صوت'],
      'Module (\\d+) \\| .*',
      'Module $1',
      { regex: true },
    );
    expect(plans.map((p) => p.paragraph)).toEqual([0, 2]);
  });
});

describe('planReplacements — literal mode', () => {
  it('treats the pattern literally by default', () => {
    const { plans } = planReplacements(['a.c and abc'], 'a.c', 'X');
    expect(plans).toHaveLength(1);
    expect(plans[0].after).toBe('X and abc');
  });

  it('is case-insensitive unless asked otherwise', () => {
    expect(planReplacements(['Hello'], 'hello', 'hi').plans).toHaveLength(1);
    expect(planReplacements(['Hello'], 'hello', 'hi', { matchCase: true }).plans).toHaveLength(0);
  });

  it('restricts to the given paragraphs', () => {
    const { plans } = planReplacements(['x', 'x', 'x'], 'x', 'y', { paragraphs: [1] });
    expect(plans.map((p) => p.paragraph)).toEqual([1]);
  });

  it('reports an empty find rather than matching everything', () => {
    expect(planReplacements(['a'], '', 'b').error).toMatch(/must not be empty/);
  });

  it('reports an invalid regex instead of throwing', () => {
    expect(planReplacements(['a'], '([', 'b', { regex: true }).error).toMatch(/invalid regex/);
  });

  it('finds nothing without inventing a plan', () => {
    expect(planReplacements(['abc'], 'zzz', 'x').plans).toEqual([]);
  });

  it('does not hang on a zero-length match', () => {
    const { plans } = planReplacements(['abc'], 'x*', 'y', { regex: true });
    expect(plans).toEqual([]);
  });
});

describe('bidi handling', () => {
  it('strips every directional mark', () => {
    expect(stripBidi(`a${RLM}b${LRM}c‪d`)).toBe('abcd');
  });

  it('can be told to respect the marks', () => {
    const text = `Module 26 ${RLM}| x`;
    expect(planReplacements([text], 'Module \\d+ \\|', 'X', { regex: true, ignoreBidiMarks: false }).plans)
      .toHaveLength(0);
  });

  it('reports oldText exactly as it sits in the document, marks included', () => {
    const text = `Module 26 ${RLM}| المقطع`;
    const { plans } = planReplacements([text], 'Module \\d+ *\\| .*', 'Module 26', { regex: true });
    expect(plans[0].oldText).toContain(RLM);
  });
});

describe('describeInvisibles', () => {
  it('names the mark the model had to hunt with charCodeAt', () => {
    expect(describeInvisibles(`ab${RLM}c`)).toEqual(['RLM (U+200F) at 2']);
  });

  it('exposes tabs, which a whitespace-collapsing preview hides', () => {
    expect(describeInvisibles('a\tb')).toEqual(['TAB at 1']);
  });

  it('names non-breaking spaces and carriage returns', () => {
    expect(describeInvisibles('a b\r')).toEqual(['NBSP at 1', 'CR at 3']);
  });

  it('is empty for plain text', () => {
    expect(describeInvisibles('plain text')).toEqual([]);
  });
});

describe('formatPlan', () => {
  it('says plainly when nothing matches', () => {
    expect(formatPlan([])).toMatch(/No matches/);
  });

  it('marks whole-paragraph versus sub-range replacements', () => {
    const { plans } = planReplacements(
      ['Module 1 | x\tModule 2 | y'],
      'Module (\\d+) \\| [^\\t]*',
      'Module $1',
      { regex: true },
    );
    const text = formatPlan(plans);
    expect(text).toContain('sub-range of paragraph');
    expect(text).toContain('2 match(es)');
  });

  it('caps the listing', () => {
    const texts = Array.from({ length: 30 }, (_, i) => `Module ${i} | x`);
    const { plans } = planReplacements(texts, 'Module (\\d+) \\| .*', 'Module $1', { regex: true });
    expect(formatPlan(plans, 5)).toContain('more not shown');
  });
});
