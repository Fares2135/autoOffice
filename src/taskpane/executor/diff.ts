// Paragraph-level diff of the document body, taken before and after an edit.
//
// Serves two readers at once: the user, who gets to see what the edit actually
// did, and the model, which gets the same summary back as the tool result
// instead of the `undefined` most generated code returns.

export interface TextDiff {
  /** Paragraphs present after but not before. */
  added: number;
  /** Paragraphs present before but not after. */
  removed: number;
  /** Unified-style lines, `-` removed and `+` added. Empty when unchanged. */
  hunks: string[];
  /** Hunks were cut for length; counts are still exact. */
  truncated: boolean;
  /** Body text is identical — the edit changed formatting only, or nothing. */
  unchanged: boolean;
  /** Document was too large to diff line by line; counts are approximate. */
  countsOnly: boolean;
}

/**
 * ponytail: O(n*m) LCS over paragraphs, capped at MAX_LCS on each side after
 * trimming the common prefix/suffix. A 300x300 table is ~90k cells, instant.
 * Beyond that we report counts only. Upgrade path if real documents hit the
 * cap: Myers diff, which is O((n+m)*d).
 */
const MAX_LCS = 300;
const MAX_HUNK_LINES = 40;
const MAX_LINE_CHARS = 160;

const splitParagraphs = (s: string): string[] =>
  s.split(/\r\n|\r|\n/).filter((line, i, all) => line !== '' || i < all.length - 1);

const clip = (s: string): string =>
  s.length <= MAX_LINE_CHARS ? s : `${s.slice(0, MAX_LINE_CHARS)}…`;

export function diffParagraphs(before: string, after: string): TextDiff {
  const a = splitParagraphs(before);
  const b = splitParagraphs(after);

  // Compare paragraphs, not raw strings: Word hands back \r while generated
  // code may write \n, and a newline-flavour change is not a text change.
  if (a.length === b.length && a.every((line, i) => line === b[i])) {
    return { added: 0, removed: 0, hunks: [], truncated: false, unchanged: true, countsOnly: false };
  }

  // Trim the identical head and tail so the LCS only sees the changed middle.
  let head = 0;
  while (head < a.length && head < b.length && a[head] === b[head]) head++;
  let tail = 0;
  while (
    tail < a.length - head &&
    tail < b.length - head &&
    a[a.length - 1 - tail] === b[b.length - 1 - tail]
  ) tail++;

  const midA = a.slice(head, a.length - tail);
  const midB = b.slice(head, b.length - tail);

  if (midA.length > MAX_LCS || midB.length > MAX_LCS) {
    return {
      added: midB.length,
      removed: midA.length,
      hunks: [],
      truncated: true,
      unchanged: false,
      countsOnly: true,
    };
  }

  // LCS lengths table.
  const n = midA.length;
  const m = midB.length;
  const lcs: number[][] = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      lcs[i][j] = midA[i] === midB[j] ? lcs[i + 1][j + 1] + 1 : Math.max(lcs[i + 1][j], lcs[i][j + 1]);
    }
  }

  const hunks: string[] = [];
  let added = 0;
  let removed = 0;
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (midA[i] === midB[j]) {
      i++;
      j++;
    } else if (lcs[i + 1][j] >= lcs[i][j + 1]) {
      hunks.push(`- ${clip(midA[i])}`);
      removed++;
      i++;
    } else {
      hunks.push(`+ ${clip(midB[j])}`);
      added++;
      j++;
    }
  }
  for (; i < n; i++) {
    hunks.push(`- ${clip(midA[i])}`);
    removed++;
  }
  for (; j < m; j++) {
    hunks.push(`+ ${clip(midB[j])}`);
    added++;
  }

  const truncated = hunks.length > MAX_HUNK_LINES;
  return {
    added,
    removed,
    hunks: truncated ? hunks.slice(0, MAX_HUNK_LINES) : hunks,
    truncated,
    unchanged: false,
    countsOnly: false,
  };
}

/** Human- and model-readable rendering of a diff. */
export function formatDiff(d: TextDiff): string {
  if (d.unchanged) return 'Document text unchanged (formatting-only edit, or no edit).';
  const header = `Document text changed: +${d.added} paragraph(s), -${d.removed} paragraph(s).`;
  if (d.countsOnly) return `${header} Document too large to list individual changes.`;
  const body = d.hunks.join('\n');
  return d.truncated ? `${header}\n${body}\n… (more changes not shown)` : `${header}\n${body}`;
}
