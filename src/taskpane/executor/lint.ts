// Static checks on generated code, run before it is offered for approval.
//
// The point is to put the risk in front of the user *before* they click approve
// — "this rewrites every paragraph in the document" is exactly what you want to
// know while you can still say no — and to hand the same warnings back to the
// model so it can narrow the edit itself.
//
// ponytail: regex-based, deliberately. A real parser would catch more, but this
// runs on model-generated office.js snippets whose shape is highly predictable,
// and a false positive here costs a sentence of explanation, not a broken edit.
// Upgrade path if the patterns get subtle: parse with acorn and walk the AST.
import type { HostKind } from '../host/context.ts';

export type WarningSeverity = 'scope' | 'performance' | 'correctness';

export interface CodeWarning {
  severity: WarningSeverity;
  /** Stable id, so the UI and tests do not depend on wording. */
  id: string;
  message: string;
}

const stripComments = (code: string): string =>
  code
    .replace(/\/\*[\s\S]*?\*\//g, ' ')
    .replace(/\/\/[^\n]*/g, ' ');

/**
 * Comments and string literals both removed, so document text that happens to
 * contain "body.clear()" cannot trigger a rule.
 */
const stripped = (code: string): string =>
  stripComments(code).replace(/(['"`])(?:\\.|(?!\1)[^\\])*\1/g, "''");

/** Does a loop body contain a context.sync() call? */
function hasSyncInLoop(code: string): boolean {
  const loop = /\b(for|while)\s*\(|\.\s*(forEach|map)\s*\(/g;
  let match: RegExpExecArray | null;
  while ((match = loop.exec(code)) !== null) {
    // Walk from the loop head to its matching close brace.
    const openIndex = code.indexOf('{', match.index);
    if (openIndex === -1) continue;
    let depth = 0;
    for (let i = openIndex; i < code.length; i++) {
      if (code[i] === '{') depth++;
      else if (code[i] === '}') {
        depth--;
        if (depth === 0) {
          if (/context\s*\.\s*sync\s*\(/.test(code.slice(openIndex, i))) return true;
          break;
        }
      }
    }
  }
  return false;
}

const WHOLE_DOCUMENT_RULES: Array<{ id: string; test: RegExp; message: string }> = [
  {
    id: 'body-clear',
    test: /\bbody\s*\.\s*clear\s*\(/,
    message: 'Clears the entire document body.',
  },
  {
    id: 'body-replace',
    test: /\bbody\s*\.\s*insert(Text|Html|Ooxml)\s*\([^)]*replace/i,
    message: 'Replaces the entire document body, which discards styles, comments, fields and tracked changes.',
  },
  {
    id: 'body-format',
    test: /\bbody\s*\.\s*(font|style|styleBuiltIn|alignment)\b\s*[.=]/,
    message: 'Applies formatting to the entire document body.',
  },
  {
    id: 'all-paragraphs-loop',
    test: /paragraphs\s*\.\s*items\s*\.\s*(forEach|map)\s*\(|for\s*\(\s*(const|let|var)\s+\w+\s+of\s+[\w.]*paragraphs\s*\.\s*items/,
    message: 'Touches every paragraph in the document.',
  },
  {
    id: 'search-all-replace',
    test: /body\s*\.\s*search\s*\([^)]*\)[\s\S]{0,200}?insert(Text|Html)\s*\([^)]*replace/i,
    message: 'Replaces every match in the whole document.',
  },
];

export function lintCode(code: string, host: HostKind): CodeWarning[] {
  const src = stripped(code);
  // load() arguments are strings, so the rules about *what is loaded* have to
  // see string contents. Comments are still gone.
  const srcWithStrings = stripComments(code);
  const warnings: CodeWarning[] = [];

  for (const rule of WHOLE_DOCUMENT_RULES) {
    if (rule.test.test(src)) {
      warnings.push({ severity: 'scope', id: rule.id, message: rule.message });
    }
  }

  if (hasSyncInLoop(src)) {
    warnings.push({
      severity: 'performance',
      id: 'sync-in-loop',
      message: 'Calls context.sync() inside a loop. Queue the work and sync once after the loop.',
    });
  }

  // Assignment can be nested (range.font.bold = true), so match the whole
  // member path, not just the last segment.
  const mutates =
    /\.\s*(insertText|insertParagraph|insertHtml|insertOoxml|insertTable|clear|delete)\s*\(/.test(src) ||
    /\.\s*(font|style|styleBuiltIn|alignment|value|values|tableDirection|sectionDirection|languageId)\s*(\.\s*\w+\s*)*=[^=]/.test(src);
  if (mutates && !/context\s*\.\s*sync\s*\(/.test(src)) {
    warnings.push({
      severity: 'correctness',
      id: 'missing-sync',
      message: 'Changes the document but never awaits context.sync(), so nothing will be applied.',
    });
  }

  const ns = host === 'word' ? 'Word' : host === 'excel' ? 'Excel' : 'PowerPoint';
  for (const other of ['Word', 'Excel', 'PowerPoint'].filter((n) => n !== ns)) {
    if (new RegExp(`\\b${other}\\s*\\.\\s*run\\s*\\(`).test(src)) {
      warnings.push({
        severity: 'correctness',
        id: 'wrong-host',
        message: `Targets ${other} but this add-in is running in ${ns}.`,
      });
    }
  }

  // APIs that do not exist, and silent no-ops. Every one of these was observed
  // costing a real turn: the model writes it, the call throws or quietly does
  // nothing, and the user waits through another round trip.
  const API_RULES: Array<{ id: string; test: RegExp; message: string }> = [
    {
      id: 'cell-width-readonly',
      test: /\.\s*(cells\s*\.\s*items\s*\[[^\]]*\]|cell)\s*\.\s*width\s*=|cells\s*\.\s*items\s*\[[^\]]*\]\s*\.\s*width\s*=/,
      message: 'TableCell.width is read-only — assigning it does nothing. Set cell.columnWidth (points) instead.',
    },
    {
      id: 'table-column-count',
      test: /\btable[\w.]*\s*\.\s*columnCount\b|items\/columnCount/,
      message: 'Word.Table has no columnCount. Take the column count from rows.items[0].cells.items.length.',
    },
    {
      id: 'paragraphs-get-item-at',
      test: /paragraphs\s*\.\s*getItemAt\s*\(/,
      message: 'ParagraphCollection has no getItemAt(). Load items and index paragraphs.items[n], or use getFirst()/getLast().',
    },
    {
      id: 'table-get-before',
      test: /\.\s*get(Before|After)OrNullObject\s*\(/,
      message: 'Word.Table has no getBeforeOrNullObject()/getAfterOrNullObject(). Index body.tables.items instead.',
    },
    {
      id: 'paragraph-replace-with-fragment',
      test: /\.\s*match\s*\([\s\S]{0,400}?insertText\s*\([^)]*(replace|Replace)/,
      message:
        'Replacing a whole paragraph with a fragment matched from it deletes everything else in ' +
        'that paragraph — index paragraphs often hold many tab-separated entries. Use the ' +
        'replace_text tool, or search inside the paragraph and replace that range.',
    },
    {
      id: 'table-values-nested',
      test: /\btables?\s*[\w.]*\s*\.\s*load\s*\(\s*['"`][^'"`]*\bvalues\b/,
      message: 'Table.values collapses nested or merged tables into one string. Load rows/items/cells/items/value, or call read_table.',
    },
  ];
  const STRING_AWARE = new Set(['table-column-count', 'table-values-nested']);
  for (const rule of API_RULES) {
    const target = STRING_AWARE.has(rule.id) ? srcWithStrings : src;
    if (rule.test.test(target)) {
      warnings.push({ severity: 'correctness', id: rule.id, message: rule.message });
    }
  }

  if (/\b(document|window)\s*\.\s*(getElementById|querySelector|createElement)/.test(src)) {
    warnings.push({
      severity: 'correctness',
      id: 'dom-access',
      message: 'Uses the DOM. Only the office.js API can change the document.',
    });
  }

  return warnings;
}

/** One line per warning, for the model's tool result. */
export function formatWarnings(warnings: CodeWarning[]): string {
  if (warnings.length === 0) return '';
  const lines = warnings.map((w) => `- [${w.severity}] ${w.message}`);
  return `Static check before running:\n${lines.join('\n')}`;
}

/** True when the code is wide enough that the user should look twice. */
export function hasScopeWarning(warnings: CodeWarning[]): boolean {
  return warnings.some((w) => w.severity === 'scope');
}


/**
 * Per-turn memory of scripts already run, so an identical re-run can be
 * answered from the previous result instead of executing again. Observed in the
 * wild: the same "dump every table" script four times in one task, each costing
 * an approval and a round trip.
 */
const runHistory = new Map<string, string>();

/** Normalise away formatting-only differences before comparing. */
export function scriptKey(code: string): string {
  return stripped(code).replace(/\s+/g, ' ').trim();
}

export function rememberRun(code: string, result: string): void {
  runHistory.set(scriptKey(code), result);
}

export function previousRun(code: string): string | undefined {
  return runHistory.get(scriptKey(code));
}

export function clearRunHistory(): void {
  runHistory.clear();
}
