// src/taskpane/skills/index.ts
import type { HostKind } from '../host/context.ts';
import { WORD_SKILLS, WORD_SKILL_NAMES } from './word/index.ts';
import { EXCEL_SKILLS, EXCEL_SKILL_NAMES } from './excel/index.ts';
import { POWERPOINT_SKILLS, POWERPOINT_SKILL_NAMES } from './powerpoint/index.ts';

const TABLES: Record<HostKind, Record<string, string>> = {
  word: WORD_SKILLS,
  excel: EXCEL_SKILLS,
  powerpoint: POWERPOINT_SKILLS,
};

const NAMES: Record<HostKind, readonly string[]> = {
  word: WORD_SKILL_NAMES,
  excel: EXCEL_SKILL_NAMES,
  powerpoint: POWERPOINT_SKILL_NAMES,
};

export function listSkills(host: HostKind): readonly string[] {
  return NAMES[host];
}

export function lookupSkill(host: HostKind, name: string): string {
  const table = TABLES[host];
  const content = table[name];
  if (!content) {
    const available = listSkills(host).join(', ');
    return `Skill "${name}" not found for host "${host}". Available: ${available}`;
  }
  return content;
}

/** Look several domains up at once, so one round-trip covers a whole edit. */
export function lookupSkills(host: HostKind, names: readonly string[]): string {
  if (names.length === 0) return lookupSkill(host, '');
  return names.map((n) => `## Skill: ${n}\n\n${lookupSkill(host, n)}`).join('\n\n---\n\n');
}

/**
 * Pure: first prose sentence of a skill document, i.e. the line after the H1
 * that is not itself a heading. Derived rather than hand-maintained so the
 * index can't drift from the docs.
 */
export function summarize(markdown: string, max = 110): string {
  const lines = markdown.split('\n').map((l) => l.trim());

  // Markdown wraps sentences across lines, so collect a whole paragraph —
  // tracking fences, or the first line of a code block reads as prose.
  const paragraph: string[] = [];
  let inFence = false;
  for (const l of lines) {
    if (l.startsWith('```')) {
      inFence = !inFence;
      if (paragraph.length) break;
      continue;
    }
    if (inFence || l.startsWith('#')) {
      if (paragraph.length) break;
      continue;
    }
    if (l === '') {
      if (paragraph.length) break;
      continue;
    }
    paragraph.push(l);
  }
  if (paragraph.length === 0) return '';

  const sentence = paragraph.join(' ').split(/(?<=\.)\s/)[0];
  return sentence.length <= max ? sentence : `${sentence.slice(0, max).trimEnd()}…`;
}

/** One `name — summary` line per skill, for the system prompt. */
export function describeSkills(host: HostKind): string {
  const table = TABLES[host];
  return listSkills(host)
    .map((name) => {
      const summary = summarize(table[name] ?? '');
      return summary ? `- ${name} — ${summary}` : `- ${name}`;
    })
    .join('\n');
}
