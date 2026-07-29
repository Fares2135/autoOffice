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
