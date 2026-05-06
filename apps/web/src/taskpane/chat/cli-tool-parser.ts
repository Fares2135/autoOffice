export type CliSegment =
  | { type: 'text'; content: string }
  | { type: 'lookup_skill'; name: string }
  | { type: 'execute_code'; code: string; index: number };

const LOOKUP_RE = /<lookup_skill>([\s\S]*?)<\/lookup_skill>/g;
const EXECUTE_RE = /<execute_code>([\s\S]*?)<\/execute_code>/g;
const ANY_TOOL_RE = /<(?:lookup_skill|execute_code)[\s\S]*?<\/(?:lookup_skill|execute_code)>/;

export function hasCliTools(text: string): boolean {
  return ANY_TOOL_RE.test(text);
}

export function parseCliSegments(text: string): CliSegment[] {
  // Collect all tag spans with their positions.
  type Span = { start: number; end: number; segment: CliSegment };
  const spans: Span[] = [];
  let execIndex = 0;

  for (const m of text.matchAll(new RegExp(LOOKUP_RE.source, 'g'))) {
    spans.push({
      start: m.index!,
      end: m.index! + m[0].length,
      segment: { type: 'lookup_skill', name: m[1]!.trim() },
    });
  }
  for (const m of text.matchAll(new RegExp(EXECUTE_RE.source, 'g'))) {
    spans.push({
      start: m.index!,
      end: m.index! + m[0].length,
      segment: { type: 'execute_code', code: m[1]!.trim(), index: execIndex++ },
    });
  }
  spans.sort((a, b) => a.start - b.start);

  const out: CliSegment[] = [];
  let cursor = 0;
  for (const { start, end, segment } of spans) {
    if (start > cursor) {
      out.push({ type: 'text', content: text.slice(cursor, start) });
    }
    out.push(segment);
    cursor = end;
  }
  if (cursor < text.length) {
    out.push({ type: 'text', content: text.slice(cursor) });
  }
  return out;
}

export function buildToolResults(
  segments: CliSegment[],
  results: Map<string, { success: boolean; output?: unknown; error?: string }>,
): string {
  const parts: string[] = [];
  for (const seg of segments) {
    if (seg.type === 'lookup_skill') {
      const r = results.get(`lookup_skill:${seg.name}`);
      if (r) {
        parts.push(
          `<lookup_skill_result name="${seg.name}">\n${r.output ?? ''}\n</lookup_skill_result>`,
        );
      }
    } else if (seg.type === 'execute_code') {
      const r = results.get(`execute_code:${seg.index}`);
      if (r) {
        parts.push(
          `<execute_code_result index="${seg.index}">\n${JSON.stringify(r.success ? { success: true, output: r.output } : { success: false, error: r.error })}\n</execute_code_result>`,
        );
      }
    }
  }
  return `<tool_results>\n${parts.join('\n')}\n</tool_results>`;
}
