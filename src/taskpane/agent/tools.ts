// src/taskpane/agent/tools.ts
import { tool, jsonSchema } from 'ai';
import { lookupSkills, listSkills } from '../skills/index.ts';
import {
  inspectDocument, findText, readParagraphs,
  readComments, readTrackedChanges, readHeadersFooters, getSelection,
} from '../executor/inspect.ts';
import { getFormatting, getStyles, readTable, restoreFormatting } from '../executor/formatting.ts';
import type { HostKind } from '../host/context.ts';

const hostName = (host: HostKind) =>
  host === 'word' ? 'Microsoft Word' : host === 'excel' ? 'Microsoft Excel' : 'Microsoft PowerPoint';

export function makeLookupSkillTool(host: HostKind) {
  const skills = listSkills(host);
  // One factory per turn, so this set is the turn's history. Re-fetching a
  // skill already in context wastes a step and tokens for nothing.
  const alreadyFetched = new Set<string>();
  return tool({
    description:
      `Fetch office.js API documentation for one or more domains in ${hostName(host)}. ` +
      `Call this before writing code to get the correct API patterns, types, and examples. ` +
      `Pass every domain the edit will touch in a single call. ` +
      `Available domains: ${skills.join(', ')}.`,
    inputSchema: jsonSchema<{ names: string[] }>({
      type: 'object',
      properties: {
        names: {
          type: 'array',
          items: { type: 'string', enum: skills as unknown as string[] },
          minItems: 1,
          description: 'Domains to fetch documentation for',
        },
      },
      required: ['names'],
      additionalProperties: false,
    }),
    execute: async ({ names }) => {
      const fresh = names.filter((n) => !alreadyFetched.has(n));
      const repeats = names.filter((n) => alreadyFetched.has(n));
      for (const n of fresh) alreadyFetched.add(n);

      if (fresh.length === 0) {
        return `Already provided earlier in this conversation: ${repeats.join(', ')}. ` +
          `Scroll up rather than fetching again.`;
      }
      const body = lookupSkills(host, fresh);
      return repeats.length
        ? `${body}\n\n---\n\nSkipped, already provided earlier: ${repeats.join(', ')}.`
        : body;
    },
  });
}

export function makeInspectDocumentTool(host: HostKind) {
  return tool({
    description:
      `Read the structure of the open ${hostName(host)} file without changing anything. ` +
      (host === 'word'
        ? 'Returns paragraph and character counts, the heading outline, the styles actually in use, ' +
          'table sizes, list and section counts, whether the text contains right-to-left script, ' +
          'the change-tracking mode, and per-section reading direction where the client supports it. '
        : 'Returns the top-level structure of the file. ') +
      'Call this before editing anything you have not already inspected — it needs no user ' +
      'approval and is cheaper than guessing wrong.',
    inputSchema: jsonSchema<Record<string, never>>({
      type: 'object',
      properties: {},
      additionalProperties: false,
    }),
    execute: async () => {
      try {
        return JSON.stringify(await inspectDocument(host), null, 2);
      } catch (err) {
        return `Inspection failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    },
  });
}

export function makeFindTextTool(host: HostKind) {
  return tool({
    description:
      'Find where text appears in the document. Read-only, no user approval, and far cheaper ' +
      'than generating a script to search. Returns each matching paragraph with its index, a ' +
      'preview and its style. ' +
      'Indexes are 0-based and line up with body.paragraphs.items[i] as of this call; inserting ' +
      'or deleting paragraphs shifts every later index, so edit from the highest index down or ' +
      'search again afterwards. ' +
      'Typical chain: find_text to locate → read_paragraphs to read the full text → execute_code ' +
      'against those indexes only. See the targeting skill for the exact edit patterns.',
    inputSchema: jsonSchema<{ query: string; matchCase?: boolean; wholeWord?: boolean }>({
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Text to look for' },
        matchCase: { type: 'boolean', description: 'Case-sensitive match (default false)' },
        wholeWord: { type: 'boolean', description: 'Match whole words only (default false)' },
      },
      required: ['query'],
      additionalProperties: false,
    }),
    execute: async ({ query, matchCase, wholeWord }) => {
      try {
        return JSON.stringify(await findText(host, query, { matchCase, wholeWord }), null, 2);
      } catch (err) {
        return `Search failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    },
  });
}

export function makeReadParagraphsTool(host: HostKind) {
  return tool({
    description:
      'Read the full text of a range of paragraphs by index, inclusive. Read-only, no user ' +
      'approval. Use this to read a section you located with inspect_document or find_text ' +
      'instead of generating a script that loads and returns text.',
    inputSchema: jsonSchema<{ from: number; to: number }>({
      type: 'object',
      properties: {
        from: { type: 'number', description: 'First paragraph index (0-based)' },
        to: { type: 'number', description: 'Last paragraph index, inclusive' },
      },
      required: ['from', 'to'],
      additionalProperties: false,
    }),
    execute: async ({ from, to }) => {
      try {
        return JSON.stringify(await readParagraphs(host, from, to), null, 2);
      } catch (err) {
        return `Read failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    },
  });
}

export function makeGetFormattingTool(host: HostKind) {
  return tool({
    description:
      'Read the exact formatting of a paragraph range: style, alignment, list membership, and ' +
      'font name, size, bold, italic, underline, colour and highlight. Read-only, no approval. ' +
      'Use it before changing formatting, and to answer questions about how something looks. ' +
      'A null value means the property is inherited from the style rather than set directly.',
    inputSchema: jsonSchema<{ from: number; to: number }>({
      type: 'object',
      properties: {
        from: { type: 'number', description: 'First paragraph index (0-based)' },
        to: { type: 'number', description: 'Last paragraph index, inclusive' },
      },
      required: ['from', 'to'],
      additionalProperties: false,
    }),
    execute: async ({ from, to }) => {
      try {
        return JSON.stringify(await getFormatting(host, from, to), null, 2);
      } catch (err) {
        return `Read failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    },
  });
}

export function makeGetStylesTool(host: HostKind) {
  return tool({
    description:
      'List the style names this document actually has. Read-only, no approval. Call it before ' +
      'applying a named style: on a localised Word install the built-in names are translated, so ' +
      '"Heading 2" may not exist and guessing it throws. Use a name from this list verbatim.',
    inputSchema: jsonSchema<Record<string, never>>({
      type: 'object',
      properties: {},
      additionalProperties: false,
    }),
    execute: async () => {
      try {
        return JSON.stringify(await getStyles(host), null, 2);
      } catch (err) {
        return `Read failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    },
  });
}

export function makeReadTableTool(host: HostKind) {
  return tool({
    description:
      'Read a table\'s cells as a grid of strings, by table index from inspect_document. ' +
      'Read-only, no approval. Use it instead of writing a script to load table values.',
    inputSchema: jsonSchema<{ index: number }>({
      type: 'object',
      properties: {
        index: { type: 'number', description: 'Table index (0-based), from inspect_document' },
      },
      required: ['index'],
      additionalProperties: false,
    }),
    execute: async ({ index }) => {
      try {
        return JSON.stringify(await readTable(host, index), null, 2);
      } catch (err) {
        return `Read failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    },
  });
}

/**
 * The repair tool. A formatting checkpoint is recorded automatically before
 * every execute_code, so "put it back the way it was" restores the recorded
 * values instead of guessing a default.
 */
export function makeRevertFormattingTool(
  host: HostKind,
  requestApproval: (summary: string) => Promise<boolean>,
) {
  return tool({
    description:
      'Undo formatting changes by restoring the checkpoint recorded just before the last edit. ' +
      'Use this whenever the user says a formatting change was wrong or asks for it to be put ' +
      'back — never guess the previous value, and never assume black, Calibri or 11pt was the ' +
      'original. Pass specific paragraph indexes to repair only those, or omit them to repair ' +
      'every paragraph whose formatting drifted from the checkpoint. Properties that were ' +
      'inherited from a style rather than set directly are reported as skipped, because office.js ' +
      'cannot re-inherit a value.',
    inputSchema: jsonSchema<{ paragraphs?: number[] }>({
      type: 'object',
      properties: {
        paragraphs: {
          type: 'array',
          items: { type: 'number' },
          description: 'Paragraph indexes to restore. Omit to restore everything that drifted.',
        },
      },
      additionalProperties: false,
    }),
    execute: async ({ paragraphs }) => {
      const scope = paragraphs?.length
        ? `paragraph(s) ${paragraphs.join(', ')}`
        : 'every paragraph whose formatting changed';
      const approved = await requestApproval(
        `Restore formatting of ${scope} to the state recorded before the last edit.`,
      );
      if (!approved) return 'User declined the formatting restore.';
      try {
        return JSON.stringify(await restoreFormatting(host, paragraphs), null, 2);
      } catch (err) {
        return `Restore failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    },
  });
}

/** Read-only tool factories for the collaboration surfaces of a document. */
export function makeReadCommentsTool(host: HostKind) {
  return tool({
    description:
      'Read the comments in the document with the text each one is anchored to, and whether it is ' +
      'resolved. Read-only, no approval. Use it when the user refers to comments or asks you to ' +
      'act on review feedback.',
    inputSchema: jsonSchema<Record<string, never>>({ type: 'object', properties: {}, additionalProperties: false }),
    execute: async () => {
      try {
        return JSON.stringify(await readComments(host), null, 2);
      } catch (err) {
        return `Read failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    },
  });
}

export function makeReadTrackedChangesTool(host: HostKind) {
  return tool({
    description:
      'Read the tracked changes still pending in the document: author, type, date and text. ' +
      'Read-only, no approval. Check this before rewriting text someone else is still reviewing — ' +
      'an edit on top of pending revisions is hard for them to untangle.',
    inputSchema: jsonSchema<Record<string, never>>({ type: 'object', properties: {}, additionalProperties: false }),
    execute: async () => {
      try {
        return JSON.stringify(await readTrackedChanges(host), null, 2);
      } catch (err) {
        return `Read failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    },
  });
}

export function makeReadHeadersFootersTool(host: HostKind) {
  return tool({
    description:
      'Read the primary header and footer text of every section. Read-only, no approval. ' +
      'Headers and footers are not part of the body, so body searches never find their text.',
    inputSchema: jsonSchema<Record<string, never>>({ type: 'object', properties: {}, additionalProperties: false }),
    execute: async () => {
      try {
        return JSON.stringify(await readHeadersFooters(host), null, 2);
      } catch (err) {
        return `Read failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    },
  });
}

/**
 * Fresh selection state. The context attached to a message is a snapshot from
 * send time; if the user moves the caret while the turn is running, or says
 * "now this one", the model needs to look again rather than trust the snapshot.
 */
export function makeReadSelectionTool(host: HostKind) {
  return tool({
    description:
      'Read what the user currently has selected: the text, its body paragraph index(es), style, ' +
      'alignment, list membership, the table/row/cell it sits in, the nearest heading above it, ' +
      'and whether it contains right-to-left script. Read-only, no approval. ' +
      'A snapshot is already attached to the user message; call this only when the selection may ' +
      'have moved since, or when that snapshot was missing what you need. ' +
      'When the selection is empty the caret position is still meaningful — that is where "here" is.',
    inputSchema: jsonSchema<Record<string, never>>({ type: 'object', properties: {}, additionalProperties: false }),
    execute: async () => {
      try {
        const sel = await getSelection(host);
        return sel ? JSON.stringify(sel, null, 2) : 'No selection information is available in this host.';
      } catch (err) {
        return `Read failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    },
  });
}
