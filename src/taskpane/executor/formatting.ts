// Formatting capture, restore, and the document's own vocabulary.
//
// office.js keeps no history. When the model sets font.color the previous value
// is gone, so "put it back the way it was" becomes a guess — and the guess is
// always black, which is wrong whenever the text was any other colour or
// inherited its colour from a style. The fix is to record the formatting
// *before* each edit and restore from that record.
import type { HostKind } from '../host/context.ts';

const MAX_SNAPSHOT_PARAGRAPHS = 1000;
const MAX_READ_PARAGRAPHS = 200;
const MAX_TABLE_CELLS = 400;
const MAX_STYLES = 60;

export interface ParagraphFormat {
  index: number;
  style: string;
  alignment: string | null;
  isListItem: boolean;
  font: {
    name: string | null;
    size: number | null;
    bold: boolean | null;
    italic: boolean | null;
    underline: string | null;
    color: string | null;
    highlightColor: string | null;
  };
}

export interface FormattingSnapshot {
  takenAt: number;
  paragraphs: ParagraphFormat[];
  /** Set when the document was too large to snapshot. */
  skippedReason?: string;
}

// One checkpoint, replaced before every edit. A deeper history would need a UI
// to pick from; the reported problem is only ever "undo what you just did".
// ponytail: single slot. Make it a stack if users start asking to go further back.
let lastSnapshot: FormattingSnapshot | null = null;

export function getLastSnapshot(): FormattingSnapshot | null {
  return lastSnapshot;
}

export function clearSnapshot(): void {
  lastSnapshot = null;
}

/** Test seam. */
export function setSnapshot(snapshot: FormattingSnapshot | null): void {
  lastSnapshot = snapshot;
}

/** Pure: the properties that differ between a recorded and a current format. */
export function changedProps(before: ParagraphFormat, after: ParagraphFormat): string[] {
  const out: string[] = [];
  if (before.style !== after.style) out.push('style');
  if (before.alignment !== after.alignment) out.push('alignment');
  for (const key of Object.keys(before.font) as Array<keyof ParagraphFormat['font']>) {
    if (before.font[key] !== after.font[key]) out.push(`font.${key}`);
  }
  return out;
}

/**
 * Pure: which properties a restore cannot honour. A null recorded value means
 * "inherited from the style, no direct formatting" — office.js has no way to
 * re-inherit, so we report it instead of silently writing a wrong literal
 * (the black-instead-of-original bug, one level down).
 */
export function unrestorableProps(before: ParagraphFormat): string[] {
  return (Object.keys(before.font) as Array<keyof ParagraphFormat['font']>)
    .filter((key) => before.font[key] === null)
    .map((key) => `font.${key}`);
}

function readParagraphFormat(p: Word.Paragraph, index: number): ParagraphFormat {
  return {
    index,
    style: p.style,
    alignment: (p.alignment as unknown as string) ?? null,
    isListItem: p.isListItem,
    font: {
      name: p.font.name ?? null,
      size: p.font.size ?? null,
      bold: p.font.bold ?? null,
      italic: p.font.italic ?? null,
      underline: (p.font.underline as unknown as string) ?? null,
      color: p.font.color ?? null,
      highlightColor: p.font.highlightColor ?? null,
    },
  };
}

const PARAGRAPH_LOAD = [
  'items/style',
  'items/alignment',
  'items/isListItem',
  'items/font/name',
  'items/font/size',
  'items/font/bold',
  'items/font/italic',
  'items/font/underline',
  'items/font/color',
  'items/font/highlightColor',
].join(',');

/** Records current formatting. Called automatically before every edit. */
export async function captureFormatting(host: HostKind): Promise<FormattingSnapshot> {
  if (host !== 'word') {
    const snapshot = { takenAt: Date.now(), paragraphs: [], skippedReason: `not supported in ${host}` };
    lastSnapshot = snapshot;
    return snapshot;
  }
  try {
    const snapshot = await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load(PARAGRAPH_LOAD);
      await context.sync();
      if (paragraphs.items.length > MAX_SNAPSHOT_PARAGRAPHS) {
        return {
          takenAt: Date.now(),
          paragraphs: [],
          skippedReason: `document has ${paragraphs.items.length} paragraphs, over the ${MAX_SNAPSHOT_PARAGRAPHS} limit`,
        };
      }
      return {
        takenAt: Date.now(),
        paragraphs: paragraphs.items.map(readParagraphFormat),
      };
    });
    lastSnapshot = snapshot;
    return snapshot;
  } catch (err) {
    const snapshot = {
      takenAt: Date.now(),
      paragraphs: [],
      skippedReason: err instanceof Error ? err.message : String(err),
    };
    lastSnapshot = snapshot;
    return snapshot;
  }
}

export interface RestoreReport {
  restored: Array<{ paragraph: number; props: string[] }>;
  unchanged: number;
  skipped: Array<{ paragraph: number; props: string[]; reason: string }>;
  outOfRange: number[];
}

/**
 * Re-applies the recorded formatting. Without `only` it repairs every paragraph
 * whose formatting drifted from the checkpoint.
 */
export async function restoreFormatting(
  host: HostKind,
  only?: number[],
): Promise<RestoreReport | { error: string }> {
  if (host !== 'word') return { error: `revert_formatting is only available in Word, not ${host}.` };
  const snapshot = lastSnapshot;
  if (!snapshot) return { error: 'No formatting checkpoint exists yet. One is recorded before each edit.' };
  if (snapshot.skippedReason) {
    return { error: `No usable checkpoint: ${snapshot.skippedReason}.` };
  }

  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load(PARAGRAPH_LOAD);
    await context.sync();

    const report: RestoreReport = { restored: [], unchanged: 0, skipped: [], outOfRange: [] };
    const wanted = only && only.length > 0
      ? snapshot.paragraphs.filter((p) => only.includes(p.index))
      : snapshot.paragraphs;

    if (only) {
      for (const index of only) {
        if (!snapshot.paragraphs.some((p) => p.index === index)) report.outOfRange.push(index);
      }
    }

    for (const before of wanted) {
      const live = paragraphs.items[before.index];
      if (!live) {
        report.outOfRange.push(before.index);
        continue;
      }
      const after = readParagraphFormat(live, before.index);
      const drifted = changedProps(before, after);
      if (drifted.length === 0) {
        report.unchanged++;
        continue;
      }

      const unrestorable = unrestorableProps(before).filter((p) => drifted.includes(p));
      const applied = drifted.filter((p) => !unrestorable.includes(p));

      if (applied.includes('style')) live.style = before.style;
      if (applied.includes('alignment') && before.alignment) {
        live.alignment = before.alignment as unknown as Word.Alignment;
      }
      if (applied.includes('font.name')) live.font.name = before.font.name!;
      if (applied.includes('font.size')) live.font.size = before.font.size!;
      if (applied.includes('font.bold')) live.font.bold = before.font.bold!;
      if (applied.includes('font.italic')) live.font.italic = before.font.italic!;
      if (applied.includes('font.underline')) {
        live.font.underline = before.font.underline as unknown as Word.UnderlineType;
      }
      if (applied.includes('font.color')) live.font.color = before.font.color!;
      if (applied.includes('font.highlightColor')) live.font.highlightColor = before.font.highlightColor!;

      if (applied.length) report.restored.push({ paragraph: before.index, props: applied });
      if (unrestorable.length) {
        report.skipped.push({
          paragraph: before.index,
          props: unrestorable,
          reason: 'was inherited from the style, not set directly — office.js cannot re-inherit a value',
        });
      }
    }

    await context.sync();
    return report;
  });
}

/** Read-only formatting of a paragraph range, so the model can look instead of guess. */
export async function getFormatting(
  host: HostKind,
  from: number,
  to: number,
): Promise<{ paragraphs: ParagraphFormat[]; total: number } | { error: string }> {
  if (host !== 'word') return { error: `get_formatting is only available in Word, not ${host}.` };

  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load(PARAGRAPH_LOAD);
    await context.sync();

    const total = paragraphs.items.length;
    const start = Math.max(0, Math.min(from, total));
    const end = Math.min(total, Math.max(start, to) + 1, start + MAX_READ_PARAGRAPHS);
    return {
      total,
      paragraphs: paragraphs.items.slice(start, end).map((p, i) => readParagraphFormat(p, start + i)),
    };
  });
}

/**
 * The style names this document actually has. Matters most on a localised
 * Word: an Arabic install names its heading style in Arabic, so "Heading 2"
 * simply does not exist and the model's guess throws.
 */
export async function getStyles(
  host: HostKind,
): Promise<{ styles: Array<{ name: string; builtIn: boolean; type: string }>; truncated: boolean } | { error: string }> {
  if (host !== 'word') return { error: `get_styles is only available in Word, not ${host}.` };

  return Word.run(async (context) => {
    // getStyles needs WordApi 1.5; fall back to the styles in use.
    try {
      const styles = context.document.getStyles();
      styles.load('items/nameLocal,items/builtIn,items/type');
      await context.sync();
      const all = styles.items.map((s) => ({
        name: s.nameLocal,
        builtIn: s.builtIn,
        type: String(s.type),
      }));
      return { styles: all.slice(0, MAX_STYLES), truncated: all.length > MAX_STYLES };
    } catch {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/style');
      await context.sync();
      const inUse = [...new Set(paragraphs.items.map((p) => p.style).filter(Boolean))];
      return {
        styles: inUse.slice(0, MAX_STYLES).map((name) => ({ name, builtIn: false, type: 'in-use' })),
        truncated: inUse.length > MAX_STYLES,
      };
    }
  });
}

/** Pure: caps a table grid so a huge table cannot flood the context. */
export function capGrid(values: string[][], maxCells = MAX_TABLE_CELLS): { grid: string[][]; truncated: boolean } {
  const out: string[][] = [];
  let cells = 0;
  for (const row of values) {
    if (cells + row.length > maxCells) return { grid: out, truncated: true };
    out.push(row);
    cells += row.length;
  }
  return { grid: out, truncated: false };
}

/** Read-only table contents as a grid. */
export async function readTable(
  host: HostKind,
  index: number,
): Promise<{ table: number; rows: number; columns: number; values: string[][]; truncated: boolean } | { error: string }> {
  if (host !== 'word') return { error: `read_table is only available in Word, not ${host}.` };

  return Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load('items/values,items/rowCount');
    await context.sync();

    const table = tables.items[index];
    if (!table) {
      return { error: `No table at index ${index}. The document has ${tables.items.length}.` };
    }
    const capped = capGrid(table.values ?? []);
    return {
      table: index,
      rows: table.rowCount,
      columns: capped.grid[0]?.length ?? 0,
      values: capped.grid,
      truncated: capped.truncated,
    };
  });
}
