import { describe, it, expect } from 'vitest';
import { probeCapabilities, formatCapabilities, officeProbe } from './capabilities.ts';

describe('probeCapabilities', () => {
  it('reports the highest supported version per family', () => {
    const supported = new Set(['WordApi 1.1', 'WordApi 1.2', 'WordApi 1.3', 'WordApiDesktop 1.1']);
    const caps = probeCapabilities('word', (n, v) => supported.has(`${n} ${v}`));
    expect(caps.WordApi).toBe('1.3');
    expect(caps.WordApiDesktop).toBe('1.1');
    expect(caps.WordApiHiddenDocument).toBeNull();
  });

  it('returns null for every family when nothing is supported', () => {
    const caps = probeCapabilities('word', () => false);
    expect(Object.values(caps).every((v) => v === null)).toBe(true);
  });

  it('probes the host-specific families only', () => {
    const asked: string[] = [];
    probeCapabilities('excel', (n) => { asked.push(n); return false; });
    expect(new Set(asked)).toEqual(new Set(['ExcelApi']));
  });

  it('treats a non-cumulative client as supporting the highest true version', () => {
    // Some clients answer true for a high set but false for a lower one.
    const caps = probeCapabilities('powerpoint', (_n, v) => v === '1.5');
    expect(caps.PowerPointApi).toBe('1.5');
  });
});

describe('officeProbe', () => {
  it('is false rather than throwing when Office is absent', () => {
    expect(officeProbe('WordApi', '1.1')).toBe(false);
  });
});

describe('formatCapabilities', () => {
  it('names unavailable families explicitly', () => {
    const text = formatCapabilities({ WordApi: '1.9', WordApiDesktop: null });
    expect(text).toContain('- WordApi: up to 1.9');
    expect(text).toContain('- WordApiDesktop: not available on this client');
  });
});
