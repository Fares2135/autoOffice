import { describe, it, expect } from 'vitest';
import { capList, distinct, preview } from './inspect.ts';

describe('capList', () => {
  it('passes short lists through untouched', () => {
    expect(capList([1, 2, 3], 5)).toEqual({ items: [1, 2, 3], truncated: false });
  });

  it('caps long lists and flags the cut', () => {
    expect(capList([1, 2, 3, 4], 2)).toEqual({ items: [1, 2], truncated: true });
  });

  it('does not flag a list that exactly fills the cap', () => {
    expect(capList([1, 2], 2).truncated).toBe(false);
  });
});

describe('distinct', () => {
  it('keeps first-seen order and drops empties', () => {
    expect(distinct(['Normal', 'Heading 1', 'Normal', '', 'Heading 2']))
      .toEqual(['Normal', 'Heading 1', 'Heading 2']);
  });
});

describe('preview', () => {
  it('collapses whitespace to a single line', () => {
    expect(preview('a\r\n  b\tc ')).toBe('a b c');
  });

  it('truncates past the limit with an ellipsis', () => {
    expect(preview('abcdef', 3)).toBe('abc…');
  });

  it('leaves Arabic text in logical order', () => {
    expect(preview('  التقرير   Q3  ')).toBe('التقرير Q3');
  });
});

import { nearestHeading, resolveIndexes, formatSelection, isPartialSelection } from './inspect.ts';

describe('nearestHeading', () => {
  const doc = [
    { style: 'Heading 1', text: 'Chapter One' },
    { style: 'Normal', text: 'body a' },
    { style: 'Heading 2', text: 'Section 1.1' },
    { style: 'Normal', text: 'body b' },
    { style: 'Normal', text: 'body c' },
  ];

  it('finds the heading above the given paragraph', () => {
    expect(nearestHeading(doc, 4)).toEqual({ paragraph: 2, style: 'Heading 2', text: 'Section 1.1' });
  });

  it('returns the heading itself when the cursor is on it', () => {
    expect(nearestHeading(doc, 2)!.paragraph).toBe(2);
  });

  it('walks further up when the nearer section has none', () => {
    expect(nearestHeading(doc, 1)!.paragraph).toBe(0);
  });

  it('is null in a document with no headings', () => {
    expect(nearestHeading([{ style: 'Normal', text: 'x' }], 0)).toBeNull();
  });

  it('clamps an out-of-range index instead of throwing', () => {
    expect(nearestHeading(doc, 99)!.paragraph).toBe(2);
  });
});

describe('resolveIndexes', () => {
  it('resolves unique paragraph text to its index', () => {
    expect(resolveIndexes(['a', 'b', 'c'], ['b'])).toEqual({ paragraphs: [1] });
  });

  it('resolves a multi-paragraph selection', () => {
    expect(resolveIndexes(['a', 'b', 'c'], ['b', 'c'])).toEqual({ paragraphs: [1, 2] });
  });

  it('reports candidates instead of guessing when the text repeats', () => {
    expect(resolveIndexes(['x', 'dup', 'y', 'dup'], ['dup']))
      .toEqual({ paragraphs: [], candidates: [1, 3] });
  });

  it('returns nothing for text that is not in the body', () => {
    expect(resolveIndexes(['a'], ['zzz'])).toEqual({ paragraphs: [] });
  });
});

describe('formatSelection', () => {
  const base = {
    empty: false, text: 'hello', partial: false, paragraphText: 'hello',
    font: null, paragraphs: [3], style: 'Normal',
    alignment: 'Left', isListItem: false, rtl: false,
    section: null, table: null,
  };

  it('is null when there is no selection at all', () => {
    expect(formatSelection(null)).toBeNull();
  });

  it('states the paragraph index so the model can target it', () => {
    expect(formatSelection(base)).toContain('paragraph index: 3');
  });

  it('says the caret is a caret, not an empty selection to ignore', () => {
    expect(formatSelection({ ...base, empty: true, text: '' }))
      .toContain('cursor position, nothing selected');
  });

  it('names the table, row and cell for "this column" style requests', () => {
    expect(formatSelection({ ...base, table: { index: 1, row: 2, cell: 0 } }))
      .toContain('inside table 1, row 2, cell 0');
  });

  it('names the enclosing heading for "this section"', () => {
    expect(formatSelection({ ...base, section: { paragraph: 2, style: 'Heading 2', text: 'Scope' } }))
      .toContain('under heading 2 "Scope"');
  });

  it('surfaces ambiguity rather than a wrong index', () => {
    const text = formatSelection({ ...base, paragraphs: [], candidates: [4, 9] })!;
    expect(text).toContain('ambiguous');
    expect(text).toContain('4, 9');
  });

  it('flags RTL selections and list membership', () => {
    const text = formatSelection({ ...base, text: 'مرحبا', rtl: true, isListItem: true })!;
    expect(text).toContain('contains right-to-left script');
    expect(text).toContain('inside a list');
  });
});

describe('isPartialSelection', () => {
  it('is true when only part of the paragraph is selected', () => {
    expect(isPartialSelection('third quarter', 'The third quarter report is ready.')).toBe(true);
  });

  it('is false when the whole paragraph is selected', () => {
    expect(isPartialSelection('The report is ready.', 'The report is ready.')).toBe(false);
  });

  it('ignores whitespace and trailing paragraph marks', () => {
    expect(isPartialSelection('The report is ready.', '  The report   is ready.\r')).toBe(false);
  });

  it('is false for an empty selection or empty paragraph', () => {
    expect(isPartialSelection('', 'anything')).toBe(false);
    expect(isPartialSelection('anything', '')).toBe(false);
  });

  it('works on Arabic text', () => {
    expect(isPartialSelection('الربع الثالث', 'تقرير الربع الثالث جاهز')).toBe(true);
    expect(isPartialSelection('تقرير الربع الثالث جاهز', 'تقرير الربع الثالث جاهز')).toBe(false);
  });
});

describe('formatSelection — partial selections and live formatting', () => {
  const base = {
    empty: false, text: 'third quarter', partial: true,
    paragraphText: 'The third quarter report is ready.',
    font: { name: 'Calibri', size: 11, bold: false, italic: false, color: '#C00000' },
    paragraphs: [3], style: 'Normal', alignment: 'Left',
    isListItem: false, rtl: false, section: null, table: null,
  };

  it('warns loudly that the paragraph is not the target', () => {
    const text = formatSelection(base)!;
    expect(text).toContain('PARTIAL selection');
    expect(text).toMatch(/not the whole paragraph/);
    expect(text).toContain('The third quarter report is ready.');
  });

  it('reports the current formatting, so "make it like before" has a value to use', () => {
    expect(formatSelection(base)!).toContain('colour #C00000');
  });

  it('reports an ambiguous table index instead of naming the wrong table', () => {
    const text = formatSelection({
      ...base, table: { index: null, candidates: [0, 2], row: 1, cell: 0 },
    })!;
    expect(text).toMatch(/index ambiguous, candidates: 0, 2/);
  });

  it('says the index is unknown when it cannot be matched at all', () => {
    const text = formatSelection({ ...base, table: { index: null, row: null, cell: null } })!;
    expect(text).toContain('table (index unknown)');
  });
});
