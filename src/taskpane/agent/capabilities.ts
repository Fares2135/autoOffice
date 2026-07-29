// Requirement-set probing.
//
// Without this the model has no idea whether the client it is driving supports
// desktop-only APIs, so it either avoids them everywhere or proposes them on
// Word on the web, where they throw. Probing at startup turns that guess into
// a fact we can put in the system prompt.
import type { HostKind } from '../host/context.ts';

/**
 * Versions to probe per requirement-set family, low to high. Requirement sets
 * are cumulative, so the highest one that reports true is the answer. Probing
 * a version that does not exist is harmless — isSetSupported just returns
 * false — so the lists deliberately reach past today's maximum.
 */
const CANDIDATES: Record<HostKind, Record<string, readonly string[]>> = {
  word: {
    WordApi: ['1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '1.8', '1.9', '1.10'],
    WordApiDesktop: ['1.1', '1.2', '1.3', '1.4'],
    WordApiHiddenDocument: ['1.3', '1.4', '1.5'],
  },
  excel: {
    ExcelApi: [
      '1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '1.8', '1.9', '1.10',
      '1.11', '1.12', '1.13', '1.14', '1.15', '1.16', '1.17', '1.18',
    ],
  },
  powerpoint: {
    PowerPointApi: ['1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '1.8'],
  },
};

/** Highest supported version per family, or null when the family is absent. */
export type Capabilities = Record<string, string | null>;

export type SetProbe = (name: string, version: string) => boolean;

/** Reads the live Office client. Never throws — a missing Office means "nothing supported". */
export const officeProbe: SetProbe = (name, version) => {
  try {
    const off = (globalThis as any).Office;
    return off?.context?.requirements?.isSetSupported(name, version) === true;
  } catch {
    return false;
  }
};

export function probeCapabilities(host: HostKind, probe: SetProbe = officeProbe): Capabilities {
  const out: Capabilities = {};
  for (const [family, versions] of Object.entries(CANDIDATES[host])) {
    let highest: string | null = null;
    for (const v of versions) {
      if (probe(family, v)) highest = v;
    }
    out[family] = highest;
  }
  return out;
}

/**
 * One line per family for the system prompt. Absent families are stated
 * explicitly — "not available" is the fact that stops the model from
 * proposing desktop-only APIs on the web.
 */
export function formatCapabilities(caps: Capabilities): string {
  const rows = Object.entries(caps).map(([family, version]) =>
    version === null
      ? `- ${family}: not available on this client`
      : `- ${family}: up to ${version}`,
  );
  return rows.join('\n');
}
