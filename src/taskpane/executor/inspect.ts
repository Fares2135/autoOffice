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
