// Typed edits: the operations users ask for by name, done by a tool instead of
// by generated code.
//
// Every one of these was reconstructed from scratch, wrongly, in a real
// transcript. Column widths took fifteen steps and never worked, because the
// obvious API (TableCell.width, Table.columns) either does nothing or does not
// exist. A tool that does it once, correctly, and then CHECKS that Word kept the
// change, replaces the whole guessing loop.
import type { HostKind } from '../host/context.ts';
import {
  captureFormatting,
  changedProps,
  getFormatting,
  widthsOf,
  type ParagraphFormat,
} from './formatting.ts';
import { getSelection } from './inspect.ts';

/** Word measures in points. Users measure in inches. */
export const PT_PER_INCH = 72;

/** Tolerance for "Word kept the width we asked for", in points. */
const WIDTH_EPSILON = 0.5;

/**
 * Pure: which requested columns Word did not actually take.
 *
 * A write that silently does nothing is the worst failure mode there is — the
 * tool reports success, the model reports success, and the document is
 * unchanged. Auto-fit tables snap their columns back on the next layout pass,
 * and office.js says nothing about it.
 */
export function unappliedColumns(
  after: number[] | null,
  columns: number[],
  widthPt: number,
): number[] {
  if (!after) return [];
  return columns.filter((c) => Math.abs((after[c] ?? 0) - widthPt) > WIDTH_EPSILON);
}

export interface ColumnWidthReport {
  table: number;
  columns: number[];
  requested: { inches: number; points: number };
  beforePt: number[] | null;
  afterPt: number[] | null;
  rowsTouched: number;
  applied: boolean;
  note?: string;
}

/**
 * Set the width of whole table columns.
 *
 * Width lives on TableCell.columnWidth and nowhere else: TableCell.width is
 * read-only, Word.Table has no columns collection and no columnCount. Setting
 * one cell would only change that cell, so every row is written.
 */
export async function setColumnWidth(
  host: HostKind,
  table: number,
  columns: number[],
  widthPt: number,
): Promise<ColumnWidthReport | { error: string }> {
  if (host !== 'word') return { error: `set_column_width is only available in Word, not ${host}.` };
  if (!Number.isFinite(widthPt) || widthPt <= 0) return { error: 'Width must be a positive number.' };
  if (columns.length === 0) return { error: 'No columns given. Pass the 0-based column indexes to resize.' };

  return Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load('items/rowCount');
    await context.sync();

    const target = tables.items[table];
    if (!target) {
      return { error: `No table at index ${table}. The document has ${tables.items.length}.` };
    }

    target.load('rows/items/cells/items/columnWidth');
    await context.sync();

    const columnCount = target.rows.items[0]?.cells.items.length ?? 0;
    const outOfRange = columns.filter((c) => !Number.isInteger(c) || c < 0 || c >= columnCount);
    if (outOfRange.length > 0) {
      return {
        error: `Table ${table} has ${columnCount} columns, so valid indexes are 0–${columnCount - 1}. ` +
          `Out of range: ${outOfRange.join(', ')}.`,
      };
    }

    const beforePt = widthsOf(target);
    for (const row of target.rows.items) {
      for (const column of columns) {
        const cell = row.cells.items[column];
        // Merged rows can be shorter than the first row; skip rather than throw.
        if (cell) cell.columnWidth = widthPt;
      }
    }
    await context.sync();

    // Read it back. Reporting a width we merely requested is how "تم بنجاح"
    // ended up on top of a document that never changed.
    target.load('rows/items/cells/items/columnWidth');
    await context.sync();
    const afterPt = widthsOf(target);
    const missed = unappliedColumns(afterPt, columns, widthPt);

    return {
      table,
      columns,
      requested: { inches: Math.round((widthPt / PT_PER_INCH) * 1000) / 1000, points: widthPt },
      beforePt,
      afterPt,
      rowsTouched: target.rows.items.length,
      applied: missed.length === 0,
      ...(missed.length > 0
        ? {
            note:
              `Word did not keep the new width on column(s) ${missed.join(', ')}. ` +
              'That normally means the table is set to auto-fit (Table Layout → AutoFit → ' +
              'Fixed column width turns it off), or the text in a cell is wider than the width asked for. ' +
              'Tell the user this rather than retrying the same call.',
          }
        : {}),
    };
  });
}

export interface FormatProps {
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  size?: number;
  fontName?: string;
  color?: string;
  highlightColor?: string;
  alignment?: string;
  style?: string;
}

export interface FormatReport {
  target: string;
  requested: FormatProps;
  changed: Array<{ paragraph: number; props: string[] }>;
  unchanged: number[];
  note?: string;
}

/** Pure: are any properties actually being asked for? */
export function hasProps(props: FormatProps): boolean {
  return Object.values(props).some((v) => v !== undefined);
}

/**
 * Apply formatting to the selection or to specific paragraphs.
 *
 * Goes through a tool rather than a script for one reason: the scope is a
 * parameter instead of something the model has to re-derive each time. A
 * checkpoint is taken first, so revert_formatting can put back what was there —
 * the previous values are otherwise gone the moment they are overwritten.
 */
export async function applyFormatting(
  host: HostKind,
  target: { selection: true } | { paragraphs: number[] },
  props: FormatProps,
): Promise<FormatReport | { error: string }> {
  if (host !== 'word') return { error: `apply_formatting is only available in Word, not ${host}.` };
  if (!hasProps(props)) return { error: 'No formatting properties given — nothing to apply.' };

  let indexes: number[];
  let label: string;
  if ('paragraphs' in target) {
    if (target.paragraphs.length === 0) return { error: 'No paragraphs given.' };
    indexes = [...target.paragraphs].sort((a, b) => a - b);
    label = `paragraph(s) ${indexes.join(', ')}`;
  } else {
    const selection = await getSelection(host);
    if (!selection) return { error: 'No selection is available. Pass paragraph indexes instead.' };
    if (selection.empty) {
      return { error: 'Nothing is selected — the caret has no text to format. Ask the user to select the text, or pass paragraph indexes.' };
    }
    indexes = selection.paragraphs;
    label = selection.partial
      ? 'the selected range (partial — narrower than its paragraph)'
      : 'the selection';
    if (indexes.length === 0 && selection.candidates?.length) {
      return {
        error: 'The selected text appears more than once, so its paragraph index is ambiguous ' +
          `(candidates: ${selection.candidates.join(', ')}). The formatting was applied to nothing; ` +
          'confirm which paragraph is meant.',
      };
    }
  }

  // Checkpoint before touching anything, so "put it back" has something to go on.
  const snapshot = await captureFormatting(host);
  const before = new Map<number, ParagraphFormat>(snapshot.paragraphs.map((p) => [p.index, p]));

  const applied = await Word.run(async (context) => {
    // Paragraph and Range both carry a font and a style name; that is all this
    // needs, so it takes the narrow shape rather than a union of the two.
    const ranges: Array<{ font: Word.Font; style: string }> = [];
    const paragraphs: Word.Paragraph[] = [];

    if ('paragraphs' in target) {
      const all = context.document.body.paragraphs;
      all.load('items');
      await context.sync();
      const missing = indexes.filter((i) => !all.items[i]);
      if (missing.length > 0) {
        return { error: `No paragraph at index ${missing.join(', ')}. The document has ${all.items.length}.` };
      }
      for (const i of indexes) paragraphs.push(all.items[i]);
      for (const p of paragraphs) ranges.push(p);
    } else {
      const range = context.document.getSelection();
      ranges.push(range);
      // alignment is paragraph-level and Word.Range does not expose it.
      const selected = range.paragraphs;
      selected.load('items');
      await context.sync();
      paragraphs.push(...selected.items);
    }

    for (const r of ranges) {
      if (props.bold !== undefined) r.font.bold = props.bold;
      if (props.italic !== undefined) r.font.italic = props.italic;
      if (props.underline !== undefined) r.font.underline = props.underline as unknown as Word.UnderlineType;
      if (props.size !== undefined) r.font.size = props.size;
      if (props.fontName !== undefined) r.font.name = props.fontName;
      if (props.color !== undefined) r.font.color = props.color;
      if (props.highlightColor !== undefined) r.font.highlightColor = props.highlightColor;
      if (props.style !== undefined) r.style = props.style;
    }
    if (props.alignment !== undefined) {
      for (const p of paragraphs) p.alignment = props.alignment as unknown as Word.Alignment;
    }

    await context.sync();
    return { error: null };
  });
  if (applied.error) return { error: applied.error };

  // Read back rather than assume. An invented style name throws; a colour Word
  // normalises comes back different; a no-op comes back identical.
  const touched = indexes.length > 0 ? indexes : [...before.keys()];
  const from = Math.min(...touched);
  const to = Math.max(...touched);
  const after = await getFormatting(host, from, to);
  if ('error' in after) {
    return { target: label, requested: props, changed: [], unchanged: [], note: after.error };
  }

  const changed: Array<{ paragraph: number; props: string[] }> = [];
  const unchanged: number[] = [];
  for (const a of after.paragraphs) {
    if (!touched.includes(a.index)) continue;
    const b = before.get(a.index);
    const drift = b ? changedProps(b, a) : [];
    if (drift.length > 0) changed.push({ paragraph: a.index, props: drift });
    else unchanged.push(a.index);
  }

  return {
    target: label,
    requested: props,
    changed,
    unchanged,
    ...(changed.length === 0
      ? {
          note:
            'Nothing changed: the paragraphs already had these values, or the request was a no-op. ' +
            'Do not repeat the call — say so, or check the target with get_formatting.',
        }
      : {}),
  };
}
