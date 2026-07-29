// src/taskpane/agent/tools.ts
import { tool, jsonSchema } from 'ai';
import { lookupSkills, listSkills } from '../skills/index.ts';
import { inspectDocument } from '../executor/inspect.ts';
import type { HostKind } from '../host/context.ts';

const hostName = (host: HostKind) =>
  host === 'word' ? 'Microsoft Word' : host === 'excel' ? 'Microsoft Excel' : 'Microsoft PowerPoint';

export function makeLookupSkillTool(host: HostKind) {
  const skills = listSkills(host);
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
    execute: async ({ names }) => lookupSkills(host, names),
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
