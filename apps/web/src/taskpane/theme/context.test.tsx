import React from 'react';
import { describe, expect, it } from 'vitest';
import { fireEvent, render, screen } from '@testing-library/react';
import { ThemeProvider, useThemeMode } from './context';

function Probe() {
  const { mode, resolved, setMode } = useThemeMode();
  return (
    <>
      <output>{mode}:{resolved}</output>
      <button onClick={() => setMode('dark')}>dark</button>
    </>
  );
}

describe('ThemeProvider', () => {
  it('persists an explicit dark theme and applies it to the document', () => {
    localStorage.clear();
    render(<ThemeProvider><Probe /></ThemeProvider>);
    fireEvent.click(screen.getByRole('button', { name: 'dark' }));
    expect(screen.getByText('dark:dark')).toBeInTheDocument();
    expect(document.documentElement.dataset.theme).toBe('dark');
    expect(localStorage.getItem('autooffice.theme')).toBe('dark');
  });
});
