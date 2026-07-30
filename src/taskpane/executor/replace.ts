// Find and replace as a tool, with a dry run.
//
// Written after watching a plain "strip the Arabic half off every Module NN |
// ... entry" task fail through a dozen hand-written scripts. Three things went
// wrong every time, and all three are handled here rather than left to the model:
//
//   1. There was no replace tool at all, so every attempt was bespoke code.
//   2. Word paragraphs in an index table hold MANY entries separated by tabs.
//      A regex ending in `.*$` therefore swallows the whole paragraph, and
//      `paragraph.insertText(match[1], 'Replace')` then deletes every sibling
//      entry — while reporting success.
//   3. Mixed Arabic/Latin text carries invisible bidi marks (U+200E/U+200F and
//      friends), so `Module 26 ‏| ...` never matches /\d+\s*\|/.
//
// So: matching ignores bidi marks by default, replacement is planned per match
// rather than per paragraph, and a match that does not cover the whole paragraph
// is applied to the sub-range only — never by rewriting the paragraph.

/** Invisible directional formatting characters. */
const BIDI_MARKS = /[‎‏‪-‮⁦-⁩؜]/g;

/** Non-global twin, for single-character tests without shared lastIndex state. */
const IS_BIDI_MARK = /[\u200E\u200F\u202A-\u202E\u2066-\u2069\u061C]/;

export const stripBidi = (s: string): string => s.replace(BIDI_MARKS, '');

/** Human-readable names, for reporting what is actually in the text. */
const MARK_NAMES: Record<string, string> = {
  '‎': 'LRM', '‏': 'RLM', '؜': 'ALM',
  '‪': 'LRE', '‫': 'RLE', '‬': 'PDF', '‭': 'LRO', '‮': 'RLO',
  '⁦': 'LRI', '⁧': 'RLI', '⁨': 'FSI', '⁩': 'PDI',
};

/**
 * Every invisible or structural character worth knowing about, with position.
 * The model was reduced to dumping charCodeAt() in a loop to discover these.
 */
export function describeInvisibles(text: string): string[] {
  const out: string[] = [];
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (MARK_NAMES[ch]) out.push(`${MARK_NAMES[ch]} (U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}) at ${i}`);
    else if (ch === '\t') out.push(`TAB at ${i}`);
    else if (ch === '\r') out.push(`CR at ${i}`);
    else if (ch === '\v') out.push(`VT (line break) at ${i}`);
    else if (ch === ' ') out.push(`NBSP at ${i}`);
  }
  return out;
}

/**
 * Renders text with its invisible structure made visible. A preview that
 * collapses whitespace hides exactly what matters here: an index paragraph full
 * of tab-separated entries looks like one sentence, and bidi marks vanish.
 */
export function escapeInvisible(text: string, max = 400): string {
  let out = '';
  for (const ch of text) {
    if (MARK_NAMES[ch]) out += `<${MARK_NAMES[ch]}>`;
    else if (ch === '\t') out += '\\t';
    else if (ch === '\r') out += '\\r';
    else if (ch === '\n') out += '\\n';
    else if (ch === '\v') out += '<VT>';
    else if (ch === '\u00A0') out += '<NBSP>';
    else out += ch;
  }
  return out.length <= max ? out : `${out.slice(0, max)}…`;
}

export interface ReplacementPlan {
  paragraph: number;
  before: string;
  after: string;
  /** Substring being replaced, as it appears in the document. */
  oldText: string;
  newText: string;
  /** True when the match covers the whole paragraph, so a paragraph-level replace is safe. */
  wholeParagraph: boolean;
  /** How many other matches share this paragraph — tabs-in-one-paragraph indexes hit this. */
  matchesInParagraph: number;
}

export interface PlanOptions {
  regex?: boolean;
  matchCase?: boolean;
  /** Default true: match across invisible bidi marks. */
  ignoreBidiMarks?: boolean;
  /** Restrict to these paragraph indexes. */
  paragraphs?: number[];
}

function buildMatchers(find: string, opts: PlanOptions): { all: RegExp; one: RegExp } {
  const source = opts.regex ? find : find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const base = opts.matchCase ? '' : 'i';
  // Two forms of the same pattern: the global one walks the paragraph, the
  // single one resolves $1-style references inside one match.
  return { all: new RegExp(source, `${base}g`), one: new RegExp(source, base) };
}

/**
 * Pure: what a replace would do, paragraph by paragraph. Runs before anything
 * is written, which is what makes a dry run possible.
 *
 * Matching happens on a bidi-stripped copy, and offsets are mapped back to the
 * original text so the reported oldText is exactly what sits in the document.
 */
export function planReplacements(
  paragraphTexts: string[],
  find: string,
  replaceWith: string,
  opts: PlanOptions = {},
): { plans: ReplacementPlan[]; error?: string } {
  if (find === '') return { plans: [], error: 'find must not be empty' };

  let matchers: { all: RegExp; one: RegExp };
  try {
    matchers = buildMatchers(find, opts);
  } catch (err) {
    return { plans: [], error: `invalid regex: ${err instanceof Error ? err.message : String(err)}` };
  }

  const ignoreBidi = opts.ignoreBidiMarks !== false;
  const plans: ReplacementPlan[] = [];

  paragraphTexts.forEach((original, index) => {
    if (opts.paragraphs && !opts.paragraphs.includes(index)) return;

    // Map every position in the stripped copy back to the original string.
    const map: number[] = [];
    let haystack = '';
    for (let i = 0; i < original.length; i++) {
      if (ignoreBidi && IS_BIDI_MARK.test(original[i])) continue;
      map.push(i);
      haystack += original[i];
    }

    matchers.all.lastIndex = 0;
    const found: Array<{ start: number; end: number; text: string; replacement: string }> = [];
    let m: RegExpExecArray | null;
    while ((m = matchers.all.exec(haystack)) !== null) {
      if (m[0] === '') break; // zero-length match would loop forever
      const startInOriginal = map[m.index];
      const endInOriginal = (map[m.index + m[0].length - 1] ?? startInOriginal) + 1;
      found.push({
        start: startInOriginal,
        end: endInOriginal,
        text: original.slice(startInOriginal, endInOriginal),
        replacement: m[0].replace(matchers.one, replaceWith),
      });
    }

    if (found.length === 0) return;

    for (const hit of found) {
      const after = original.slice(0, hit.start) + hit.replacement + original.slice(hit.end);
      plans.push({
        paragraph: index,
        before: original,
        after,
        oldText: hit.text,
        newText: hit.replacement,
        wholeParagraph: stripBidi(hit.text).trim() === stripBidi(original).trim(),
        matchesInParagraph: found.length,
      });
    }
  });

  return { plans };
}

/** Pure: compact, reviewable rendering of a plan. */
export function formatPlan(plans: ReplacementPlan[], limit = 20): string {
  if (plans.length === 0) return 'No matches — nothing would change.';
  const shown = plans.slice(0, limit).map((p) => {
    const scope = p.wholeParagraph
      ? 'whole paragraph'
      : `sub-range of paragraph (${p.matchesInParagraph} match(es) in this paragraph)`;
    return `- paragraph ${p.paragraph} [${scope}]\n    - "${p.oldText}"\n    + "${p.newText}"`;
  });
  const head = `${plans.length} match(es) in ${new Set(plans.map((p) => p.paragraph)).size} paragraph(s):`;
  return plans.length > limit
    ? `${head}\n${shown.join('\n')}\n… ${plans.length - limit} more not shown`
    : `${head}\n${shown.join('\n')}`;
}

import type { HostKind } from '../host/context.ts';

export interface ReplaceReport {
  applied: Array<{ paragraph: number; oldText: string; newText: string }>;
  skipped: Array<{ paragraph: number; oldText: string; reason: string }>;
  matches: number;
  paragraphsTouched: number;
}

/**
 * Reads every paragraph so a plan can be built. Split out from apply so the dry
 * run costs nothing but a read.
 */
export async function planInDocument(
  host: HostKind,
  find: string,
  replaceWith: string,
  opts: PlanOptions = {},
): Promise<{ plans: ReplacementPlan[]; error?: string }> {
  if (host !== 'word') return { plans: [], error: `replace_text is only available in Word, not ${host}.` };
  const texts = await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text');
    await context.sync();
    return paragraphs.items.map((p) => p.text);
  });
  return planReplacements(texts, find, replaceWith, opts);
}

/**
 * Applies a plan. The safety rule that the failed transcript needed: a match
 * that does not cover the whole paragraph is replaced through a search inside
 * that paragraph, so its siblings survive. If the sub-range cannot be isolated,
 * the paragraph is SKIPPED and reported — never rewritten wholesale, which is
 * what silently deleted fifteen index entries at a time.
 */
export async function applyReplacements(
  host: HostKind,
  plans: ReplacementPlan[],
): Promise<ReplaceReport | { error: string }> {
  if (host !== 'word') return { error: `replace_text is only available in Word, not ${host}.` };

  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text');
    await context.sync();

    const report: ReplaceReport = { applied: [], skipped: [], matches: plans.length, paragraphsTouched: 0 };

    // Highest index first: replacing text never shifts paragraph indexes, but
    // multiple matches inside one paragraph do shift each other's offsets, and
    // searching for the literal old text sidesteps that entirely.
    for (const plan of plans) {
      const paragraph = paragraphs.items[plan.paragraph];
      if (!paragraph) {
        report.skipped.push({ paragraph: plan.paragraph, oldText: plan.oldText, reason: 'paragraph no longer exists' });
        continue;
      }

      if (plan.wholeParagraph) {
        paragraph.insertText(plan.newText, Word.InsertLocation.replace);
        report.applied.push({ paragraph: plan.paragraph, oldText: plan.oldText, newText: plan.newText });
        continue;
      }

      // Sub-range: locate the exact old text inside this paragraph only.
      const hits = paragraph.search(plan.oldText, { matchCase: true });
      hits.load('items');
      await context.sync();

      if (hits.items.length === 0) {
        report.skipped.push({
          paragraph: plan.paragraph,
          oldText: plan.oldText,
          reason: 'could not locate the text inside the paragraph (invisible characters or Word search limits). ' +
            'Not rewriting the whole paragraph, because that would delete the rest of it.',
        });
        continue;
      }

      hits.items[0].insertText(plan.newText, Word.InsertLocation.replace);
      report.applied.push({ paragraph: plan.paragraph, oldText: plan.oldText, newText: plan.newText });
    }

    await context.sync();
    report.paragraphsTouched = new Set(report.applied.map((a) => a.paragraph)).size;
    return report;
  });
}
