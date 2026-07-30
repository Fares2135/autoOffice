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

import {
  nearestHeading, resolveIndexes, formatSelection, isPartialSelection,
  formatSelectionTable, summariseCells, ptToIn, type SelectionTable,
} from './inspect.ts';

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

  it('names the table and the cells for "this column" style requests', () => {
    const text = formatSelection({ ...base, table: tbl({ index: 1, cells: [{ row: 2, column: 0 }] }) })!;
    expect(text).toContain('inside table 1');
    expect(text).toContain('row(s) 2 × column(s) 0');
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
      ...base, table: tbl({ index: null, candidates: [0, 2], cells: [{ row: 1, column: 0 }] }),
    })!;
    expect(text).toMatch(/index ambiguous, candidates: 0, 2/);
  });

  it('says the index is unknown when it cannot be matched at all', () => {
    const text = formatSelection({ ...base, table: tbl({ index: null }) })!;
    expect(text).toContain('table (index unknown)');
  });
});

/** A SelectionTable with the boring fields filled in. */
function tbl(over: Partial<SelectionTable> & { index: number | null }): SelectionTable {
  const cells = over.cells ?? [];
  return {
    rows: 7,
    columns: 6,
    row: cells[0]?.row ?? null,
    cell: cells[0]?.column ?? null,
    cells,
    ...summariseCells(cells, over.rows ?? 7),
    columnWidthsPt: null,
    ...over,
  };
}

describe('summariseCells', () => {
  it('folds the covered cells into distinct rows and columns', () => {
    const cells = [
      { row: 0, column: 1 }, { row: 0, column: 2 },
      { row: 1, column: 1 }, { row: 1, column: 2 },
    ];
    expect(summariseCells(cells, 2)).toEqual({
      selectedRows: [0, 1], selectedColumns: [1, 2], wholeColumns: true,
    });
  });

  it('sorts ascending however the cells arrived', () => {
    const cells = [{ row: 2, column: 5 }, { row: 0, column: 1 }, { row: 2, column: 1 }];
    expect(summariseCells(cells, 5).selectedColumns).toEqual([1, 5]);
    expect(summariseCells(cells, 5).selectedRows).toEqual([0, 2]);
  });

  it('is not whole columns when only some rows are covered', () => {
    expect(summariseCells([{ row: 0, column: 0 }], 7).wholeColumns).toBe(false);
  });

  it('handles no cells at all', () => {
    expect(summariseCells([], 3)).toEqual({ selectedRows: [], selectedColumns: [], wholeColumns: false });
  });
});

// The reported failure: five cells selected across a row, and the model could
// not tell which columns they were, so it counted letters in the selected text
// and guessed five adjacent columns starting from the first cell.
describe('formatSelectionTable — "change these columns to 0.6\\""', () => {
  const selected = [1, 2, 3, 4, 5].flatMap((column) =>
    [0, 1, 2, 3, 4, 5, 6].map((row) => ({ row, column })),
  );
  const table = tbl({ index: 2, cells: selected, columnWidthsPt: [86.4, 64.8, 64.8, 64.8, 64.8, 64.8] });

  it('names the table index without another lookup', () => {
    expect(formatSelectionTable(table)).toContain('inside table 2');
  });

  it('lists the selected columns, so nothing has to be inferred from the text', () => {
    expect(formatSelectionTable(table)).toContain('COLUMN(S) 1, 2, 3, 4, 5');
  });

  it('says the columns are covered in full', () => {
    expect(formatSelectionTable(table)).toContain('all 7 rows');
  });

  it('gives the current widths in inches as well as points', () => {
    const text = formatSelectionTable(table);
    expect(text).toContain('col 0: 1.2"');
    expect(text).toContain('86.4pt');
  });

  it('reports the grid size', () => {
    expect(formatSelectionTable(table)).toContain('7 rows × 6 columns');
  });

  it('tells the model to apply partial selections to whole columns', () => {
    const partial = tbl({
      index: 2,
      cells: [{ row: 3, column: 1 }, { row: 3, column: 2 }],
    });
    expect(formatSelectionTable(partial)).toContain('apply to every row');
  });
});

describe('ptToIn', () => {
  it('converts points to inches at two decimals', () => {
    expect(ptToIn(43.2)).toBe(0.6);
    expect(ptToIn(72)).toBe(1);
  });
});
