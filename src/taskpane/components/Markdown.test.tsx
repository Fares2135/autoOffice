import React from 'react';
import { describe, it, expect } from 'vitest';
import { render, screen } from '@testing-library/react';
import { Markdown } from './Markdown.tsx';

describe('Markdown', () => {
  it('renders emphasis as elements, not literal asterisks', () => {
    const { container } = render(<Markdown>{'This is **bold** text'}</Markdown>);
    expect(container.querySelector('strong')?.textContent).toBe('bold');
    expect(container.textContent).not.toContain('**');
  });

  it('renders lists, headings and inline code', () => {
    const { container } = render(
      <Markdown>{'## Title\n\n- one\n- two\n\nUse `range.load()`'}</Markdown>,
    );
    expect(container.querySelector('h2')?.textContent).toBe('Title');
    expect(container.querySelectorAll('li')).toHaveLength(2);
    expect(container.querySelector('code')?.textContent).toBe('range.load()');
  });

  it('renders GFM tables inside a scroll container', () => {
    const { container } = render(
      <Markdown>{'| a | b |\n| - | - |\n| 1 | 2 |'}</Markdown>,
    );
    const table = container.querySelector('table');
    expect(table).not.toBeNull();
    expect(container.querySelectorAll('th')).toHaveLength(2);
    // The table is wrapped so wide output scrolls instead of stretching the pane.
    expect(table!.parentElement!.tagName).toBe('DIV');
  });

  it('renders fenced code blocks', () => {
    const { container } = render(
      <Markdown>{'```js\nconst a = 1;\n```'}</Markdown>,
    );
    expect(container.querySelector('pre code')?.textContent).toContain('const a = 1;');
  });

  it('does not render raw HTML from model output', () => {
    const { container } = render(<Markdown>{'<img src=x onerror=alert(1)>'}</Markdown>);
    expect(container.querySelector('img')).toBeNull();
  });

  it('opens links in a new tab without leaking the referrer', () => {
    render(<Markdown>{'[docs](https://example.com)'}</Markdown>);
    const link = screen.getByRole('link');
    expect(link).toHaveAttribute('target', '_blank');
    expect(link).toHaveAttribute('rel', 'noopener noreferrer');
  });

  it('lets each block pick its own direction so Arabic and English mix', () => {
    const { container } = render(<Markdown>{'مرحبا **بك**'}</Markdown>);
    expect(container.firstElementChild).toHaveAttribute('dir', 'auto');
    expect(container.querySelector('strong')?.textContent).toBe('بك');
  });
});
