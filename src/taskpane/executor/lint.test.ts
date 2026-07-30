import { describe, it, expect, beforeEach } from 'vitest';
import { lintCode, formatWarnings, hasScopeWarning, noOpNote, looksLikeTextWrite } from './lint.ts';

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

import { scriptKey, rememberRun, previousRun, clearRunHistory } from './lint.ts';

describe('lintCode — APIs that cost real turns', () => {
  it('catches the silent no-op: assigning TableCell.width', () => {
    const w = lintCode('row.cells.items[0].width = 43.2; await context.sync();', 'word');
    expect(w.map(x => x.id)).toContain('cell-width-readonly');
    expect(w.find(x => x.id === 'cell-width-readonly')!.message).toContain('columnWidth');
  });

  it('catches Table.columnCount, which does not exist', () => {
    expect(ids('tables.load("items/values, items/columnCount"); await context.sync();'))
      .toContain('table-column-count');
    expect(ids('const c = table.columnCount;')).toContain('table-column-count');
  });

  it('catches paragraphs.getItemAt(), which does not exist', () => {
    expect(ids('const p = context.document.body.paragraphs.getItemAt(516);'))
      .toContain('paragraphs-get-item-at');
  });

  it('catches table.getBeforeOrNullObject(), which does not exist', () => {
    expect(ids('const t = table.getBeforeOrNullObject();')).toContain('table-get-before');
  });

  it('warns that Table.values collapses nested tables', () => {
    expect(ids('tables.load("items/values"); await context.sync();')).toContain('table-values-nested');
  });

  it('does not warn about the correct cell-by-cell load', () => {
    expect(ids('table.load("rows/items/cells/items/value"); await context.sync();'))
      .not.toContain('table-values-nested');
  });

  it('does not warn about the correct columnWidth assignment', () => {
    expect(ids('row.cells.items[0].columnWidth = 43.2; await context.sync();'))
      .not.toContain('cell-width-readonly');
  });
});

describe('run history', () => {
  beforeEach(() => clearRunHistory());

  it('treats whitespace and comment differences as the same script', () => {
    expect(scriptKey('a();  // note\n b();')).toBe(scriptKey('a();\nb();'));
  });

  it('recalls the previous result for an identical script', () => {
    rememberRun('doThing();', 'result A');
    expect(previousRun('doThing();')).toBe('result A');
    expect(previousRun(' doThing();  ')).toBe('result A');
  });

  it('does not confuse two different scripts', () => {
    rememberRun('doThing();', 'A');
    expect(previousRun('doOther();')).toBeUndefined();
  });

  it('is cleared between turns', () => {
    rememberRun('x();', 'A');
    clearRunHistory();
    expect(previousRun('x();')).toBeUndefined();
  });
});

describe('lintCode — the silent paragraph destroyer', () => {
  it('flags replacing a paragraph with a fragment matched from it', () => {
    const code = `
      const m = p.text.match(/(Module\\s+\\d+)\\s*\\|.*$/);
      if (m) { p.insertText(m[1], Word.InsertLocation.replace); }
      await context.sync();
    `;
    const w = lintCode(code, 'word');
    expect(w.map(x => x.id)).toContain('paragraph-replace-with-fragment');
    expect(w.find(x => x.id === 'paragraph-replace-with-fragment')!.message).toContain('replace_text');
  });

  it('does not flag matching without a replace', () => {
    expect(ids('const m = p.text.match(/x/); console.log(m);'))
      .not.toContain('paragraph-replace-with-fragment');
  });

  it('does not flag an append after a match', () => {
    expect(ids('const m = t.match(/x/); body.insertParagraph(m[0], Word.InsertLocation.end);'))
      .not.toContain('paragraph-replace-with-fragment');
  });
});

// Straight out of the failed transcript: five scripts, each "successful", none
// of which could have worked — Word.Table simply has no columns collection.
describe('the APIs the column-width task invented', () => {
  it('flags table.columns.items', () => {
    expect(ids('const t = sel.parentTableCell.parentRow.parentTable; console.log(t.columns.items.length);'))
      .toContain('table-columns-collection');
  });

  it('flags loading columns/items', () => {
    expect(ids('table.load("columns/items"); await context.sync();'))
      .toContain('table-columns-collection');
  });

  it('flags writing through the non-existent collection', () => {
    const w = lintCode('table.columns.items[c].width = 43.2; await context.sync();', 'word');
    expect(w.map(x => x.id)).toContain('table-columns-collection');
    expect(w.find(x => x.id === 'table-columns-collection')!.message).toContain('set_column_width');
  });

  it('does not flag the pattern that works', () => {
    expect(ids('for (const row of table.rows.items) row.cells.items[1].columnWidth = 43.2; await context.sync();'))
      .not.toContain('table-columns-collection');
  });

  it('flags asking a multi-cell selection for its single parent cell', () => {
    const w = lintCode(
      'const sel = context.document.getSelection();\nconst cell = sel.parentTableCellOrNullObject;\ncell.load("rowIndex");',
      'word',
    );
    expect(w.map(x => x.id)).toContain('selection-parent-cell-single');
    expect(w.find(x => x.id === 'selection-parent-cell-single')!.message).toContain('read_selection');
  });

  it('does not flag a paragraph asking for its own cell, which is the right way', () => {
    expect(ids('const cell = paragraphs.items[3].parentTableCellOrNullObject; cell.load("rowIndex");'))
      .not.toContain('selection-parent-cell-single');
  });
});

describe('noOpNote', () => {
  const write = 'p.insertText("x", Word.InsertLocation.replace); await context.sync();';

  it('warns when a text write left the text identical', () => {
    const note = noOpNote(write, undefined, true);
    expect(note).toContain('did nothing');
    expect(note).toContain('Do not report success');
  });

  it('stays quiet when the text did change', () => {
    expect(noOpNote(write, undefined, false)).toBe('');
  });

  it('stays quiet for a formatting-only script, which legitimately changes no text', () => {
    expect(noOpNote('range.font.bold = true; await context.sync();', undefined, true)).toBe('');
  });

  it('stays quiet for a script that returned data — that was a read', () => {
    expect(noOpNote(write, { rows: 3 }, true)).toBe('');
  });

  it('stays quiet for a script that logged', () => {
    expect(noOpNote(write, undefined, true, ['Table index: 2'])).toBe('');
  });

  it('does not treat document text mentioning insertText as a write', () => {
    expect(noOpNote('return "call insertText to replace";', undefined, true)).toBe('');
  });
});

describe('looksLikeTextWrite', () => {
  it('spots value assignment on a cell', () => {
    expect(looksLikeTextWrite('cell.value = "0.6";')).toBe(true);
  });

  it('does not count a comparison as a write', () => {
    expect(looksLikeTextWrite('if (cell.value === "x") {}')).toBe(false);
  });

  it('does not count a width change as a text write', () => {
    expect(looksLikeTextWrite('cell.columnWidth = 43.2;')).toBe(false);
  });
});
