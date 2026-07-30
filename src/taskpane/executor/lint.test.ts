import { describe, it, expect } from 'vitest';
import { lintCode, formatWarnings, hasScopeWarning } from './lint.ts';

const ids = (code: string, host: 'word' | 'excel' | 'powerpoint' = 'word') =>
  lintCode(code, host).map((w) => w.id);

describe('lintCode — scope', () => {
  it('flags clearing the whole body', () => {
    expect(ids('context.document.body.clear(); await context.sync();')).toContain('body-clear');
  });

  it('flags replacing the whole body', () => {
    expect(ids(`
      context.document.body.insertText(newText, Word.InsertLocation.replace);
      await context.sync();
    `)).toContain('body-replace');
  });

  it('flags formatting applied to the whole body', () => {
    expect(ids('context.document.body.font.color = "#FF0000"; await context.sync();'))
      .toContain('body-format');
  });

  it('flags a loop over every paragraph', () => {
    expect(ids(`
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();
      paragraphs.items.forEach(p => { p.font.bold = true; });
      await context.sync();
    `)).toContain('all-paragraphs-loop');
  });

  it('leaves a narrowly targeted edit alone', () => {
    const warnings = lintCode(`
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();
      paragraphs.items[7].font.bold = true;
      await context.sync();
    `, 'word');
    expect(warnings).toEqual([]);
  });

  it('leaves a selection edit alone', () => {
    expect(lintCode('context.document.getSelection().font.bold = true; await context.sync();', 'word'))
      .toEqual([]);
  });
});

describe('lintCode — performance', () => {
  it('flags context.sync() inside a for loop', () => {
    expect(ids(`
      for (const p of items) {
        p.font.bold = true;
        await context.sync();
      }
    `)).toContain('sync-in-loop');
  });

  it('flags context.sync() inside forEach', () => {
    expect(ids('items.forEach(async p => { p.font.bold = true; await context.sync(); });'))
      .toContain('sync-in-loop');
  });

  it('accepts a loop that syncs once afterwards', () => {
    expect(ids(`
      for (const p of items) { p.font.bold = true; }
      await context.sync();
    `)).not.toContain('sync-in-loop');
  });
});

describe('lintCode — correctness', () => {
  it('flags a mutation with no sync at all', () => {
    expect(ids('context.document.getSelection().font.bold = true;')).toContain('missing-sync');
  });

  it('does not ask a read-only script to sync-and-mutate', () => {
    expect(ids('const b = context.document.body; b.load("text"); await context.sync(); return b.text;'))
      .not.toContain('missing-sync');
  });

  it('flags the wrong host namespace', () => {
    expect(ids('Excel.run(async (context) => {});', 'word')).toContain('wrong-host');
    expect(ids('Word.run(async (context) => {});', 'word')).not.toContain('wrong-host');
  });

  it('flags DOM access', () => {
    expect(ids('document.getElementById("x").innerHTML = "y";')).toContain('dom-access');
  });
});

describe('lintCode — no false positives from text content', () => {
  it('ignores rule text inside strings', () => {
    expect(ids('range.insertText("body.clear() is dangerous", Word.InsertLocation.end); await context.sync();'))
      .not.toContain('body-clear');
  });

  it('ignores rule text inside comments', () => {
    expect(ids(`
      // do not call body.clear() here
      /* body.font.color = "red" */
      context.document.getSelection().font.bold = true;
      await context.sync();
    `)).toEqual([]);
  });

  it('does not flag Arabic replacement text as a body replace', () => {
    expect(ids('paras.items[3].insertText("التقرير النهائي", Word.InsertLocation.replace); await context.sync();'))
      .toEqual([]);
  });
});

describe('formatWarnings / hasScopeWarning', () => {
  it('is empty for clean code', () => {
    expect(formatWarnings([])).toBe('');
    expect(hasScopeWarning([])).toBe(false);
  });

  it('renders one line per warning with its severity', () => {
    const text = formatWarnings(lintCode('context.document.body.clear();', 'word'));
    expect(text).toContain('Static check before running');
    expect(text).toMatch(/\[scope\]/);
  });

  it('separates scope risk from mere performance advice', () => {
    expect(hasScopeWarning(lintCode('context.document.body.clear(); await context.sync();', 'word'))).toBe(true);
    expect(hasScopeWarning(lintCode('for (const p of items) { await context.sync(); }', 'word'))).toBe(false);
  });
});
