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

  it('forbids touching anything the user did not ask about', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toContain('Change exactly what was asked and nothing else');
    expect(p).toMatch(/narrowest range/);
  });

  it('refuses to treat a scoped request as a whole-document licence', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toMatch(/NEVER a licence to operate on the whole document/);
    expect(p).toContain('document.body');
  });

  it('names the whole-document operations to treat as destructive, per host', () => {
    expect(buildSystemPrompt('word', 'en')).toContain('body.clear()');
    expect(buildSystemPrompt('excel', 'en')).toMatch(/every sheet or slide/);
  });

  it('tells the model to report an over-broad edit instead of claiming success', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toMatch(/outside what was asked/);
    expect(p).toContain('Ctrl+Z');
  });

  it('points the model at the read-only tools instead of read-only scripts', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toContain('find_text');
    expect(p).toContain('read_paragraphs');
    expect(p).toMatch(/Never write a script whose only purpose is to read or search/);
  });

  it('forbids guessing a previous formatting value', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toContain('NEVER guess a previous value');
    expect(p).toContain('revert_formatting');
    expect(p).toMatch(/Black, Calibri and 11pt are not "the original"/);
  });

  it('tells the model to read formatting before writing it', () => {
    expect(buildSystemPrompt('word', 'en')).toContain('get_formatting');
  });

  it('forbids inventing style names on a localised Word', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toContain('get_styles');
    expect(p).toMatch(/Never invent a style name/);
  });

  it('lists every read-only tool as approval-free', () => {
    const p = buildSystemPrompt('word', 'en');
    for (const tool of ['inspect_document', 'find_text', 'read_paragraphs', 'get_formatting', 'get_styles', 'read_table']) {
      expect(p).toContain(tool);
    }
    expect(p).toMatch(/read-only and need no user approval/);
  });

  it('bans context.sync() inside loops', () => {
    const p = buildSystemPrompt('word', 'en');
    expect(p).toContain('NEVER call context.sync() inside a loop');
  });
});
