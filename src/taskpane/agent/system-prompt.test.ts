import { describe, it, expect } from 'vitest';
import { buildSystemPrompt } from './system-prompt.ts';

describe('buildSystemPrompt', () => {
  it('contains an English-locale clause naming "English"', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toMatch(/respond to the user in \*\*English\*\* \(en\)/i);
  });

  it('contains a Hebrew-locale clause naming the native name', () => {
    const p = buildSystemPrompt('word', 'he');
    expect(p).toMatch(/respond to the user in \*\*עברית\*\* \(he\)/i);
  });

  it('contains an Arabic-locale clause naming the native name', () => {
    const p = buildSystemPrompt('word', 'ar');
    expect(p).toMatch(/respond to the user in \*\*العربية\*\* \(ar\)/i);
  });

  it('keeps locale clause near the end of the prompt', () => {
    const p = buildSystemPrompt('word', 'he');
    const idx = p.toLowerCase().indexOf('respond to the user');
    expect(idx).toBeGreaterThan(p.length / 2);
  });

  it('still includes the office.js critical rules', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toContain('CRITICAL RULES for office.js code');
  });

  it('lists skills with a one-line summary, not bare names', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toMatch(/- tables — .+/);
    expect(p).toMatch(/- arabic-rtl — .+/);
  });

  it('states which requirement sets the client supports, and which it does not', () => {
    const p = buildSystemPrompt('word', 'en', { WordApi: '1.9', WordApiDesktop: null });
    expect(p).toContain('WordApi: up to 1.9');
    expect(p).toContain('WordApiDesktop: not available on this client');
  });

  it('omits the capability clause when nothing was probed', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).not.toContain('requirement sets');
  });

  it('asks for a per-item report on batch edits', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toMatch(/ok: 18, failed:/);
  });

  it('tells the model to verify the change report before claiming success', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toContain('Document text changed');
  });

  it('points the model at inspect_document before editing', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toMatch(/inspect_document first/);
  });

  it('bans context.sync() inside loops', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toContain('NEVER call context.sync() inside a loop');
  });
});
