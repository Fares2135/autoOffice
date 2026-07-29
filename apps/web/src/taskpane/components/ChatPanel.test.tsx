import React from 'react';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { ChatPanel } from './ChatPanel';
import { LanguageProvider } from '../i18n';
import { translationService } from '../i18n/service';

const baseProps = {
  host: { kind: 'word' as const, displayName: 'Word' },
  messages: [],
  status: 'ready',
  onSubmit: vi.fn(),
  onApproveCode: vi.fn(),
  onRejectCode: vi.fn(),
  onApprovalResponse: vi.fn(),
  highlightCode: (code: string) => <pre>{code}</pre>,
};

async function renderPanel(
  props: Partial<React.ComponentProps<typeof ChatPanel>> = {},
  locale: 'en' | 'ar' = 'en',
) {
  await translationService.setLocale(locale);
  return render(
    <LanguageProvider initialLocale={locale}>
      <FluentProvider theme={webLightTheme}>
        <ChatPanel {...baseProps} {...props} />
      </FluentProvider>
    </LanguageProvider>,
  );
}

describe('ChatPanel onboarding and prompts', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    Element.prototype.scrollIntoView = vi.fn();
    globalThis.requestAnimationFrame = (callback) => {
      callback(0);
      return 1;
    };
  });

  it('guides an unconfigured user to settings', async () => {
    const onOpenSettings = vi.fn();
    await renderPanel({ noProvider: true, onOpenSettings });
    await waitFor(() => expect(screen.getByRole('heading', { name: /Connect your AI workspace/i })).toBeInTheDocument());
    fireEvent.click(screen.getByRole('button', { name: /Open settings/i }));
    expect(onOpenSettings).toHaveBeenCalledOnce();
    expect(screen.getByRole('button', { name: 'Send' })).toBeDisabled();
  });

  it('fills the composer from a Word prompt card', async () => {
    await renderPanel();
    const card = await screen.findByRole('button', { name: /Polish this document/i });
    fireEvent.click(card);
    expect(screen.getByRole('textbox')).toHaveValue('Polish this document');
  });

  it('renders Arabic onboarding in RTL while keeping editable text direction automatic', async () => {
    await renderPanel({ noProvider: true }, 'ar');
    await waitFor(() => expect(document.documentElement).toHaveAttribute('dir', 'rtl'));
    expect(screen.getByRole('textbox')).toHaveAttribute('dir', 'auto');
    expect(screen.getByRole('heading').textContent).toContain('اربط مساحة عمل');
  });
});
