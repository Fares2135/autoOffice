// Read-only document inspection.
//
// The model used to have to write code and wait for approval just to look at
// the document, so it usually guessed instead. This reads the structure
// directly — no code generation, no approval, nothing written.
import type { HostKind } from '../host/context.ts';
import { officeProbe, type SetProbe } from '../agent/capabilities.ts';

const MAX_HEADINGS = 40;
const MAX_STYLES = 20;
const MAX_TABLES = 20;
const SELECTION_CHARS = 300;

export interface WordOutline {
  host: 'word';
  paragraphs: number;
  characters: number;
  headings: Array<{ style: string; text: string }>;
  headingsTruncated?: true;
  stylesInUse: string[];
  stylesTruncated?: true;
  tables: Array<{ rows: number }>;
  tablesTruncated?: true;
  listParagraphs: number;
  sections: number;
  /** Body text contains Arabic or Hebrew characters. */
  hasBidiText: boolean;
  /** Only present when the client exposes it. */
  changeTrackingMode?: string;
  sectionDirections?: string[];
  notes?: string[];
}

export interface SimpleOutline {
  host: 'excel' | 'powerpoint';
  sheets?: string[];
  slides?: number;
  notes?: string[];
}

export type Outline = WordOutline | SimpleOutline;

// Hebrew, Arabic, Arabic Supplement/Extended-A and both presentation-form blocks.
const BIDI = /[\u0590-\u05FF\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/;

/** Pure: caps a list and reports whether anything was dropped. */
export function capList<T>(items: T[], max: number): { items: T[]; truncated: boolean } {
  return items.length <= max
    ? { items, truncated: false }
    : { items: items.slice(0, max), truncated: true };
}

/** Pure: distinct values in first-seen order. */
export function distinct(values: string[]): string[] {
  return [...new Set(values.filter(Boolean))];
}

/** Pure: single-line, length-capped preview of a text run. */
export function preview(text: string, max = SELECTION_CHARS): string {
  const oneLine = text.replace(/\s+/g, ' ').trim();
  return oneLine.length <= max ? oneLine : `${oneLine.slice(0, max)}…`;
}

export async function inspectDocument(host: HostKind, probe: SetProbe = officeProbe): Promise<Outline> {
  if (host === 'word') return inspectWord(probe);
  if (host === 'excel') return inspectExcel();
  return inspectPowerPoint();
}

async function inspectWord(probe: SetProbe): Promise<WordOutline> {
  return Word.run(async (context) => {
    const notes: string[] = [];
    const body = context.document.body;
    body.load('text');

    const paragraphs = body.paragraphs;
    paragraphs.load('items/style,items/text,items/isListItem');

    const tables = body.tables;
    tables.load('items/rowCount');

    const sections = context.document.sections;
    sections.load('items');

    await context.sync();

    const paraItems = paragraphs.items;
    const headingItems = paraItems
      .filter((p) => typeof p.style === 'string' && /^heading/i.test(p.style))
      .map((p) => ({ style: p.style, text: preview(p.text, 120) }));
    const headings = capList(headingItems, MAX_HEADINGS);
    const styles = capList(distinct(paraItems.map((p) => p.style)), MAX_STYLES);
    const tableRows = capList(tables.items.map((t) => ({ rows: t.rowCount })), MAX_TABLES);

    const out: WordOutline = {
      host: 'word',
      paragraphs: paraItems.length,
      characters: body.text.length,
      headings: headings.items,
      stylesInUse: styles.items,
      tables: tableRows.items,
      listParagraphs: paraItems.filter((p) => p.isListItem).length,
      sections: sections.items.length,
      hasBidiText: BIDI.test(body.text),
    };
    if (headings.truncated) out.headingsTruncated = true;
    if (styles.truncated) out.stylesTruncated = true;
    if (tableRows.truncated) out.tablesTruncated = true;

    // Anything version-gated is read in its own sync so one unsupported
    // property cannot take the whole inspection down with it.
    if (probe('WordApi', '1.4')) {
      try {
        const doc = context.document;
        doc.load('changeTrackingMode');
        await context.sync();
        out.changeTrackingMode = doc.changeTrackingMode;
      } catch {
        notes.push('changeTrackingMode could not be read.');
      }
    }

    if (probe('WordApiDesktop', '1.3')) {
      try {
        const setups = sections.items.map((s) => {
          const ps = s.pageSetup;
          ps.load('sectionDirection');
          return ps;
        });
        await context.sync();
        out.sectionDirections = setups.map((ps) => String(ps.sectionDirection));
      } catch {
        notes.push('sectionDirection could not be read.');
      }
    } else {
      notes.push('WordApiDesktop 1.3 unavailable: section/table direction and bidi font properties cannot be set on this client.');
    }

    if (notes.length) out.notes = notes;
    return out;
  });
}

async function inspectExcel(): Promise<SimpleOutline> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    return { host: 'excel', sheets: sheets.items.map((s) => s.name) };
  });
}

async function inspectPowerPoint(): Promise<SimpleOutline> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    return { host: 'powerpoint', slides: slides.items.length };
  });
}

const MAX_HITS = 50;
const MAX_READ_PARAGRAPHS = 100;

export interface SearchHit {
  paragraph: number;
  text: string;
  style: string;
}

/**
 * Read-only search. Replaces "write a script, wait for approval, run it, read
 * the log" for the single most common question the model has about a document:
 * where does this text appear?
 */
export async function findText(
  host: HostKind,
  query: string,
  opts: { matchCase?: boolean; wholeWord?: boolean } = {},
): Promise<{ query: string; hits: SearchHit[]; total: number; truncated: boolean } | { error: string }> {
  if (host !== 'word') return { error: `find_text is only available in Word, not ${host}.` };
  if (!query) return { error: 'query must not be empty.' };

  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text,items/style');
    await context.sync();

    const needle = opts.matchCase ? query : query.toLowerCase();
    const hits: SearchHit[] = [];
    paragraphs.items.forEach((p, index) => {
      const haystack = opts.matchCase ? p.text : p.text.toLowerCase();
      const found = opts.wholeWord
        ? new RegExp(`\\b${escapeRegExp(needle)}\\b`).test(haystack)
        : haystack.includes(needle);
      if (found) hits.push({ paragraph: index, text: preview(p.text, 200), style: p.style });
    });

    const capped = capList(hits, MAX_HITS);
    return { query, hits: capped.items, total: hits.length, truncated: capped.truncated };
  });
}

/** Pure: escapes a user string for use inside a RegExp. */
export function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Read-only paragraph slice, so the model can read a section it already
 * located without generating code for it.
 */
export async function readParagraphs(
  host: HostKind,
  from: number,
  to: number,
): Promise<{ from: number; to: number; total: number; paragraphs: Array<{ index: number; text: string; style: string }> } | { error: string }> {
  if (host !== 'word') return { error: `read_paragraphs is only available in Word, not ${host}.` };

  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text,items/style');
    await context.sync();

    const total = paragraphs.items.length;
    const start = Math.max(0, Math.min(from, total));
    const end = Math.min(total, Math.max(start, to) + 1, start + MAX_READ_PARAGRAPHS);
    return {
      from: start,
      to: end - 1,
      total,
      paragraphs: paragraphs.items.slice(start, end).map((p, i) => ({
        index: start + i,
        text: p.text,
        style: p.style,
      })),
    };
  });
}

/**
 * Body text for the before/after diff. Returns null for hosts where a text
 * snapshot is not meaningful, which skips the diff entirely.
 */
export async function getBodyText(host: HostKind): Promise<string | null> {
  if (host !== 'word') return null;
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      body.load('text');
      await context.sync();
      return body.text;
    });
  } catch {
    return null;
  }
}

/**
 * What the user currently has selected. "Make this bold" is meaningless
 * without it, and the model would otherwise burn a turn asking.
 */
export async function getSelectionContext(host: HostKind): Promise<string | null> {
  if (host !== 'word') return null;
  try {
    return await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load('text,style');
      const paragraphs = range.paragraphs;
      paragraphs.load('items/alignment');
      await context.sync();

      if (!range.text) return null;
      const alignment = paragraphs.items[0]?.alignment;
      const parts = [
        `text: "${preview(range.text)}"`,
        range.style ? `style: ${range.style}` : '',
        alignment ? `alignment: ${alignment}` : '',
        BIDI.test(range.text) ? 'contains right-to-left script' : '',
      ].filter(Boolean);
      return parts.join(', ');
    });
  } catch {
    return null;
  }
}
