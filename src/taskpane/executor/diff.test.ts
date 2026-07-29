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
