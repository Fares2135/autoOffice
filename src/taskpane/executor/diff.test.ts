import { describe, it, expect } from 'vitest';
import { diffParagraphs, formatDiff } from './diff.ts';

describe('diffParagraphs', () => {
  it('reports identical text as unchanged', () => {
    const d = diffParagraphs('a\rb', 'a\rb');
    expect(d.unchanged).toBe(true);
    expect(d.hunks).toEqual([]);
    expect(formatDiff(d)).toMatch(/unchanged/i);
  });

  it('counts a replaced paragraph as one added and one removed', () => {
    const d = diffParagraphs('intro\rold body\rend', 'intro\rnew body\rend');
    expect(d.added).toBe(1);
    expect(d.removed).toBe(1);
    expect(d.hunks).toEqual(['- old body', '+ new body']);
  });

  it('detects a pure insertion without touching the surrounding paragraphs', () => {
    const d = diffParagraphs('a\rb', 'a\rmiddle\rb');
    expect(d).toMatchObject({ added: 1, removed: 0 });
    expect(d.hunks).toEqual(['+ middle']);
  });

  it('detects a pure deletion', () => {
    const d = diffParagraphs('a\rgone\rb', 'a\rb');
    expect(d).toMatchObject({ added: 0, removed: 1 });
    expect(d.hunks).toEqual(['- gone']);
  });

  it('handles Arabic text without mangling it', () => {
    const d = diffParagraphs('مرحبا\rالتقرير', 'مرحبا\rالتقرير Q3');
    expect(d.hunks).toEqual(['- التقرير', '+ التقرير Q3']);
  });

  it('caps the listed hunks but keeps the counts exact', () => {
    const before = Array.from({ length: 60 }, (_, i) => `old ${i}`).join('\r');
    const after = Array.from({ length: 60 }, (_, i) => `new ${i}`).join('\r');
    const d = diffParagraphs(before, after);
    expect(d.added).toBe(60);
    expect(d.removed).toBe(60);
    expect(d.truncated).toBe(true);
    expect(d.hunks.length).toBeLessThanOrEqual(40);
    expect(formatDiff(d)).toContain('more changes not shown');
  });

  it('falls back to counts only past the LCS ceiling', () => {
    const before = Array.from({ length: 400 }, (_, i) => `a${i}`).join('\r');
    const after = Array.from({ length: 400 }, (_, i) => `b${i}`).join('\r');
    const d = diffParagraphs(before, after);
    expect(d.countsOnly).toBe(true);
    expect(d.hunks).toEqual([]);
    expect(formatDiff(d)).toMatch(/too large/i);
  });

  it('clips very long paragraphs in the hunk output', () => {
    const d = diffParagraphs('x'.repeat(500), 'y'.repeat(500));
    for (const h of d.hunks) expect(h.length).toBeLessThan(200);
  });

  it('treats CRLF, CR and LF as the same paragraph break', () => {
    expect(diffParagraphs('a\r\nb', 'a\nb').unchanged).toBe(true);
    expect(diffParagraphs('a\rb', 'a\nb').unchanged).toBe(true);
  });
});

import { attachContext, setLastEditTargets, getLastEditTargets } from './diff.ts';

describe('changedIndexes', () => {
  it('points at where the new text ended up', () => {
    const d = diffParagraphs('a\rold\rc', 'a\rnew\rc');
    expect(d.changedIndexes).toEqual([1]);
  });

  it('covers an insertion', () => {
    expect(diffParagraphs('a\rb', 'a\rmid\rb').changedIndexes).toEqual([1]);
  });

  it('is empty for a pure deletion', () => {
    expect(diffParagraphs('a\rgone\rb', 'a\rb').changedIndexes).toEqual([]);
  });

  it('is empty when nothing changed', () => {
    expect(diffParagraphs('a\rb', 'a\rb').changedIndexes).toEqual([]);
  });

  it('lists every changed paragraph', () => {
    const d = diffParagraphs('a\rx\ry\rb', 'a\rp\rq\rb');
    expect(d.changedIndexes).toEqual([1, 2]);
  });
});

describe('attachContext', () => {
  it('rides the selection note along with the message', () => {
    const out = attachContext('make this bold', 'selected text: "hello"; paragraph index: 3');
    expect(out).toContain('make this bold');
    expect(out).toContain('[Current selection — selected text: "hello"; paragraph index: 3]');
  });

  it('leaves the message untouched when there is nothing to attach', () => {
    expect(attachContext('hello', null)).toBe('hello');
  });

  it('adds the previous-edit referent', () => {
    const out = attachContext('make it bigger', null, [4, 5]);
    expect(out).toMatch(/previous edit changed paragraph\(s\) 4, 5/);
    expect(out).toMatch(/"it" and "that" refer here/);
  });

  it('carries both notes at once', () => {
    const out = attachContext('do it here', 'cursor position, nothing selected', [7]);
    expect(out).toContain('[Current selection —');
    expect(out).toContain('previous edit changed');
  });
});

describe('last-edit slot', () => {
  it('remembers the most recent target only', () => {
    setLastEditTargets([1]);
    setLastEditTargets([9, 10]);
    expect(getLastEditTargets()).toEqual([9, 10]);
  });
});
