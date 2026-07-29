import { describe, it, expect } from 'vitest';
import { systemPromptForHost } from './system-prompt';

describe('systemPromptForHost', () => {
  it('mentions Word for word host', () => {
    expect(systemPromptForHost('word')).toMatch(/Word/);
  });
  it('mentions Excel for excel host', () => {
    expect(systemPromptForHost('excel')).toMatch(/Excel/);
  });
  it('mentions PowerPoint for powerpoint host', () => {
    expect(systemPromptForHost('powerpoint')).toMatch(/PowerPoint/);
  });

  it('protects Arabic and mixed-direction Word content', () => {
    const prompt = systemPromptForHost('word');
    expect(prompt).toContain('never reverse strings manually');
    expect(prompt).toContain('lookup_skill("arabic-rtl")');
    expect(prompt).toContain('embedded Latin text, numbers, URLs');
  });

  it('instructs the assistant to reply in the user language', () => {
    expect(systemPromptForHost('word')).toContain('Reply in the language used by the user');
  });
});
