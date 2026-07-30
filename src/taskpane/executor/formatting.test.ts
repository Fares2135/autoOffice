import { describe, it, expect, beforeEach } from 'vitest';
import {
  changedProps,
  unrestorableProps,
  capGrid,
  getLastSnapshot,
  setSnapshot,
  clearSnapshot,
  restoreFormatting,
  type ParagraphFormat,
} from './formatting.ts';

const fmt = (over: Partial<ParagraphFormat['font']> = {}, rest: Partial<ParagraphFormat> = {}): ParagraphFormat => ({
  index: 0,
  style: 'Normal',
  alignment: 'Left',
  isListItem: false,
  font: {
    name: 'Calibri',
    size: 11,
    bold: false,
    italic: false,
    underline: 'None',
    color: '#000000',
    highlightColor: null,
    ...over,
  },
  ...rest,
});

describe('changedProps', () => {
  it('finds nothing when formatting is identical', () => {
    expect(changedProps(fmt(), fmt())).toEqual([]);
  });

  it('names the exact font property that drifted', () => {
    expect(changedProps(fmt({ color: '#FF0000' }), fmt({ color: '#0000FF' })))
      .toEqual(['font.color']);
  });

  it('reports the real reported bug: colour changed away from a non-black original', () => {
    const before = fmt({ color: '#C00000' });
    const after = fmt({ color: '#0000FF' });
    expect(changedProps(before, after)).toEqual(['font.color']);
    // The recorded value is the original, not black.
    expect(before.font.color).toBe('#C00000');
  });

  it('catches style and alignment changes too', () => {
    expect(changedProps(fmt({}, { style: 'Normal' }), fmt({}, { style: 'Heading 1' })))
      .toContain('style');
    expect(changedProps(fmt({}, { alignment: 'Left' }), fmt({}, { alignment: 'Right' })))
      .toContain('alignment');
  });

  it('reports several drifted properties at once', () => {
    const props = changedProps(fmt(), fmt({ bold: true, size: 14 }));
    expect(props).toEqual(expect.arrayContaining(['font.bold', 'font.size']));
  });
});

describe('unrestorableProps', () => {
  it('flags inherited (null) values, which cannot be re-inherited', () => {
    expect(unrestorableProps(fmt({ highlightColor: null }))).toEqual(['font.highlightColor']);
  });

  it('is empty when everything was set directly', () => {
    expect(unrestorableProps(fmt({ highlightColor: '#FFFF00' }))).toEqual([]);
  });
});

describe('capGrid', () => {
  it('passes a small table through', () => {
    expect(capGrid([['a', 'b'], ['c', 'd']])).toEqual({ grid: [['a', 'b'], ['c', 'd']], truncated: false });
  });

  it('stops before exceeding the cell budget', () => {
    const wide = Array.from({ length: 10 }, () => ['1', '2', '3']);
    const { grid, truncated } = capGrid(wide, 9);
    expect(grid).toHaveLength(3);
    expect(truncated).toBe(true);
  });

  it('handles an empty table', () => {
    expect(capGrid([])).toEqual({ grid: [], truncated: false });
  });
});

describe('snapshot slot', () => {
  beforeEach(() => clearSnapshot());

  it('starts empty', () => {
    expect(getLastSnapshot()).toBeNull();
  });

  it('holds the most recent checkpoint only', () => {
    setSnapshot({ takenAt: 1, paragraphs: [fmt()] });
    setSnapshot({ takenAt: 2, paragraphs: [] });
    expect(getLastSnapshot()!.takenAt).toBe(2);
  });
});

describe('restoreFormatting guards', () => {
  beforeEach(() => clearSnapshot());

  it('refuses on non-Word hosts', async () => {
    expect(await restoreFormatting('excel')).toEqual({
      error: 'revert_formatting is only available in Word, not excel.',
    });
  });

  it('explains that no checkpoint exists yet', async () => {
    const result = await restoreFormatting('word');
    expect(result).toHaveProperty('error');
    expect((result as { error: string }).error).toMatch(/No formatting checkpoint/);
  });

  it('explains why a checkpoint is unusable rather than restoring nothing silently', async () => {
    setSnapshot({ takenAt: 1, paragraphs: [], skippedReason: 'document has 5000 paragraphs, over the 1000 limit' });
    const result = await restoreFormatting('word');
    expect((result as { error: string }).error).toMatch(/over the 1000 limit/);
  });
});
