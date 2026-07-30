// Read-only document inspection.
//
// The model used to have to write code and wait for approval just to look at
// the document, so it usually guessed instead. This reads the structure
// directly — no code generation, no approval, nothing written.
import type { HostKind } from '../host/context.ts';
import { officeProbe, type SetProbe } from '../agent/capabilities.ts';
import { escapeInvisible } from './replace.ts';

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
  tables: Array<{ index: number; rows: number; columns: number; firstRow: string[] }>;
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
    // Row count alone cannot tell six tables apart. Columns and the first row
    // are what let the model pick the right one without dumping every table.
    tables.load('items/rowCount,items/rows/items/cells/items/value');

    const sections = context.document.sections;
    sections.load('items');

    await context.sync();

    const paraItems = paragraphs.items;
    const headingItems = paraItems
      .filter((p) => typeof p.style === 'string' && /^heading/i.test(p.style))
      .map((p) => ({ style: p.style, text: preview(p.text, 120) }));
    const headings = capList(headingItems, MAX_HEADINGS);
    const styles = capList(distinct(paraItems.map((p) => p.style)), MAX_STYLES);
    const tableRows = capList(
      tables.items.map((t, i) => {
        const firstRowCells = t.rows.items[0]?.cells.items ?? [];
        return {
          index: i,
          rows: t.rowCount,
          columns: firstRowCells.length,
          firstRow: firstRowCells.slice(0, 8).map((c) => preview(c.value ?? '', 40)),
        };
      }),
      MAX_TABLES,
    );

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
  /** True when the hit is inside a table — a body search alone never says so. */
  inTable?: boolean;
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
      // escapeInvisible, not preview: a collapsed preview turns a tab-separated
      // index paragraph into one unreadable sentence and hides bidi marks.
      if (found) hits.push({ paragraph: index, text: escapeInvisible(p.text, 240), style: p.style });
    });

    // Table membership per hit: knowing a match sits in a table, and being able
    // to hand that paragraph index to table_for_paragraph, is what turns
    // "find the table that contains B C C B D" into two calls instead of ten
    // exploratory scripts.
    const capped = capList(hits, MAX_HITS);
    const parents = capped.items.map((hit) => {
      const parent = paragraphs.items[hit.paragraph].parentTableOrNullObject;
      parent.load('isNullObject');
      return parent;
    });
    try {
      await context.sync();
      capped.items.forEach((hit, i) => {
        if (!parents[i].isNullObject) hit.inTable = true;
      });
    } catch {
      // Older client without parentTableOrNullObject; hits still stand.
    }

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
): Promise<{ from: number; to: number; total: number; paragraphs: Array<{ index: number; text: string; visible: string; style: string }> } | { error: string }> {
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
        // Same text with tabs, line breaks and bidi marks spelled out.
        visible: escapeInvisible(p.text, 600),
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
export interface SelectionContext {
  /** True when there is a caret but no selected text. */
  empty: boolean;
  text: string;
  /**
   * True when the selection covers only part of its paragraph. Formatting the
   * whole paragraph in that case is the classic "not what I meant" edit.
   */
  partial: boolean;
  /** Text of the paragraph(s) the selection sits in, for comparison. */
  paragraphText: string;
  /** Direct formatting of the selected range itself. */
  font: {
    name: string | null;
    size: number | null;
    bold: boolean | null;
    italic: boolean | null;
    color: string | null;
  } | null;
  /** Body paragraph indexes the selection covers, when they could be resolved. */
  paragraphs: number[];
  /** Set when the same text appears more than once and the index is ambiguous. */
  candidates?: number[];
  style: string | null;
  alignment: string | null;
  isListItem: boolean;
  rtl: boolean;
  /** Nearest heading at or above the selection — what the user means by "this section". */
  section: { paragraph: number; style: string; text: string } | null;
  table: {
    index: number | null;
    candidates?: number[];
    row: number | null;
    cell: number | null;
  } | null;
}

/** Pure: nearest heading at or above `index`, i.e. the section the user is in. */
export function nearestHeading(
  paragraphs: Array<{ style: string; text: string }>,
  index: number,
): { paragraph: number; style: string; text: string } | null {
  for (let i = Math.min(index, paragraphs.length - 1); i >= 0; i--) {
    const p = paragraphs[i];
    if (p && typeof p.style === 'string' && /^heading/i.test(p.style)) {
      return { paragraph: i, style: p.style, text: preview(p.text, 120) };
    }
  }
  return null;
}

/**
 * Pure: resolve selected paragraph text to body indexes. Exact when the text is
 * unique; otherwise every candidate is reported rather than picking one, because
 * silently guessing the wrong paragraph is how an edit lands in the wrong place.
 */
export function resolveIndexes(
  bodyTexts: string[],
  selectedTexts: string[],
): { paragraphs: number[]; candidates?: number[] } {
  const paragraphs: number[] = [];
  const candidates: number[] = [];
  for (const text of selectedTexts) {
    const hits = bodyTexts.reduce<number[]>((acc, t, i) => (t === text ? [...acc, i] : acc), []);
    if (hits.length === 1) paragraphs.push(hits[0]);
    else candidates.push(...hits);
  }
  return candidates.length > 0 && paragraphs.length === 0
    ? { paragraphs, candidates: [...new Set(candidates)] }
    : { paragraphs };
}

/**
 * Pure: is the selection narrower than the paragraph(s) it lies in? Compared on
 * whitespace-collapsed text, because Word reports paragraph text with trailing
 * marks the range does not include.
 */
export function isPartialSelection(selectedText: string, paragraphText: string): boolean {
  const a = selectedText.replace(/\s+/g, ' ').trim();
  const b = paragraphText.replace(/\s+/g, ' ').trim();
  if (a === '' || b === '') return false;
  return a !== b && b.includes(a);
}

/** Pure: renders the selection for the model, or null when there is nothing useful to say. */
export function formatSelection(sel: SelectionContext | null): string | null {
  if (!sel) return null;
  const parts: string[] = [];
  parts.push(sel.empty ? 'cursor position, nothing selected' : `selected text: "${preview(sel.text)}"`);
  if (sel.partial) {
    parts.push(
      'PARTIAL selection — it covers only part of its paragraph, so edit the selected range ' +
      'itself, not the whole paragraph',
    );
    parts.push(`containing paragraph: "${preview(sel.paragraphText, 160)}"`);
  }
  if (sel.font) {
    const f = sel.font;
    const bits = [
      f.name ? `font ${f.name}` : '',
      f.size ? `${f.size}pt` : '',
      f.bold ? 'bold' : '',
      f.italic ? 'italic' : '',
      f.color ? `colour ${f.color}` : '',
    ].filter(Boolean);
    if (bits.length) parts.push(`current formatting: ${bits.join(', ')}`);
  }
  if (sel.paragraphs.length === 1) parts.push(`paragraph index: ${sel.paragraphs[0]}`);
  else if (sel.paragraphs.length > 1) parts.push(`paragraph indexes: ${sel.paragraphs.join(', ')}`);
  else if (sel.candidates?.length) parts.push(`paragraph index ambiguous, candidates: ${sel.candidates.join(', ')}`);
  if (sel.style) parts.push(`style: ${sel.style}`);
  if (sel.alignment) parts.push(`alignment: ${sel.alignment}`);
  if (sel.isListItem) parts.push('inside a list');
  if (sel.table) {
    const which = sel.table.index !== null
      ? `table ${sel.table.index}`
      : sel.table.candidates?.length
        ? `a table (index ambiguous, candidates: ${sel.table.candidates.join(', ')})`
        : 'a table (index unknown)';
    const where = sel.table.row !== null && sel.table.cell !== null
      ? `${which}, row ${sel.table.row}, cell ${sel.table.cell}`
      : which;
    parts.push(`inside ${where}`);
  }
  if (sel.section) parts.push(`under heading ${sel.section.paragraph} "${sel.section.text}"`);
  if (sel.rtl) parts.push('contains right-to-left script');
  return parts.join('; ');
}

/**
 * What the user is pointing at. "Make this bold", "add it here", "this column"
 * are all unresolvable without it — and an empty selection still matters,
 * because the caret is where "here" is.
 */
export async function getSelection(host: HostKind): Promise<SelectionContext | null> {
  if (host !== 'word') return null;

  // Built in stages, each guarded on its own. A failure in the table lookup
  // must not throw away the whole note — losing the selection entirely is far
  // worse than losing one field of it.
  let base: SelectionContext | null = null;
  try {
    base = await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load('text,style,font/name,font/size,font/bold,font/italic,font/color');

      const selected = range.paragraphs;
      selected.load('items/text,items/style,items/alignment,items/isListItem');

      const body = context.document.body.paragraphs;
      body.load('items/text,items/style');

      await context.sync();

      const first = selected.items[0];
      const paragraphText = selected.items.map((p) => p.text).join('\n');
      const resolved = resolveIndexes(
        body.items.map((p) => p.text),
        selected.items.map((p) => p.text),
      );
      const anchor = resolved.paragraphs[0] ?? resolved.candidates?.[0] ?? 0;
      const text = range.text ?? '';

      return {
        empty: !text,
        text,
        partial: isPartialSelection(text, paragraphText),
        paragraphText: preview(paragraphText, 300),
        font: {
          name: range.font?.name ?? null,
          size: range.font?.size ?? null,
          bold: range.font?.bold ?? null,
          italic: range.font?.italic ?? null,
          color: range.font?.color ?? null,
        },
        paragraphs: resolved.paragraphs,
        ...(resolved.candidates ? { candidates: resolved.candidates } : {}),
        style: range.style ?? first?.style ?? null,
        alignment: (first?.alignment as unknown as string) ?? null,
        isListItem: first?.isListItem ?? false,
        rtl: BIDI.test(text),
        section: nearestHeading(body.items.map((p) => ({ style: p.style, text: p.text })), anchor),
        table: null,
      } as SelectionContext;
    });
  } catch {
    return null;
  }

  // Table context is version-gated and optional, so it gets its own attempt.
  try {
    const table = await Word.run(async (context) => {
      const range = context.document.getSelection();
      const parent = range.parentTableOrNullObject;
      parent.load('isNullObject,values,rowCount');
      const cell = range.parentTableCellOrNullObject;
      cell.load('isNullObject,rowIndex,cellIndex');
      const tables = context.document.body.tables;
      tables.load('items/values');
      await context.sync();

      if (parent.isNullObject) return null;
      // Proxy objects are never identity-equal, so match on content and report
      // ambiguity rather than mislabelling the table.
      const key = JSON.stringify(parent.values ?? []);
      const matches = tables.items
        .map((t, i) => ({ i, key: JSON.stringify(t.values ?? []) }))
        .filter((t) => t.key === key)
        .map((t) => t.i);
      return {
        index: matches.length === 1 ? matches[0] : null,
        candidates: matches.length > 1 ? matches : undefined,
        row: cell.isNullObject ? null : cell.rowIndex,
        cell: cell.isNullObject ? null : cell.cellIndex,
      };
    });
    if (table) base.table = table;
  } catch {
    // No table context; the rest of the note still stands.
  }

  return base;
}

/** Back-compat wrapper used when attaching context to a user message. */
export async function getSelectionContext(host: HostKind): Promise<string | null> {
  return formatSelection(await getSelection(host));
}

const MAX_COMMENTS = 50;
const MAX_REVISIONS = 50;

/**
 * Comments with the text they are anchored to. Needs WordApi 1.4; returns a
 * clear reason rather than throwing when the client is older.
 */
export async function readComments(
  host: HostKind,
  probe: SetProbe = officeProbe,
): Promise<{ comments: Array<{ author: string; text: string; anchoredTo: string; resolved: boolean }>; total: number; truncated: boolean } | { error: string }> {
  if (host !== 'word') return { error: `read_comments is only available in Word, not ${host}.` };
  if (!probe('WordApi', '1.4')) return { error: 'Reading comments needs WordApi 1.4, which this client does not support.' };

  return Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load('items/authorName,items/content,items/resolved');
    const ranges = comments.load('items');
    await context.sync();

    const anchors = ranges.items.map((c) => {
      const range = c.getRange();
      range.load('text');
      return range;
    });
    await context.sync();

    const all = comments.items.map((c, i) => ({
      author: c.authorName,
      text: preview(c.content, 300),
      anchoredTo: preview(anchors[i]?.text ?? '', 120),
      resolved: c.resolved,
    }));
    const capped = capList(all, MAX_COMMENTS);
    return { comments: capped.items, total: all.length, truncated: capped.truncated };
  });
}

/**
 * Tracked changes still pending in the document. Needs WordApi 1.6.
 * Knowing these exist changes how an edit should be made, so the model should
 * check before rewriting text someone else is still reviewing.
 */
export async function readTrackedChanges(
  host: HostKind,
  probe: SetProbe = officeProbe,
): Promise<{ changes: Array<{ author: string; type: string; date: string; text: string }>; total: number; truncated: boolean } | { error: string }> {
  if (host !== 'word') return { error: `read_tracked_changes is only available in Word, not ${host}.` };
  if (!probe('WordApi', '1.6')) return { error: 'Reading tracked changes needs WordApi 1.6, which this client does not support.' };

  return Word.run(async (context) => {
    const revisions = context.document.body.getTrackedChanges();
    revisions.load('items/author,items/type,items/date,items/text');
    await context.sync();

    const all = revisions.items.map((r) => ({
      author: r.author,
      type: String(r.type),
      date: String(r.date),
      text: preview(r.text ?? '', 200),
    }));
    const capped = capList(all, MAX_REVISIONS);
    return { changes: capped.items, total: all.length, truncated: capped.truncated };
  });
}

/** Header and footer text per section — otherwise only reachable by writing a script. */
export async function readHeadersFooters(
  host: HostKind,
): Promise<{ sections: Array<{ section: number; header: string; footer: string }> } | { error: string }> {
  if (host !== 'word') return { error: `read_headers_footers is only available in Word, not ${host}.` };

  return Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load('items');
    await context.sync();

    const parts = sections.items.map((s) => {
      const header = s.getHeader(Word.HeaderFooterType.primary);
      const footer = s.getFooter(Word.HeaderFooterType.primary);
      header.load('text');
      footer.load('text');
      return { header, footer };
    });
    await context.sync();

    return {
      sections: parts.map((p, i) => ({
        section: i,
        header: preview(p.header.text, 300),
        footer: preview(p.footer.text, 300),
      })),
    };
  });
}
