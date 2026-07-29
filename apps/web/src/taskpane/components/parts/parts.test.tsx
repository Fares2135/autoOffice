import React from 'react';
import { describe, it, expect, vi, afterEach, beforeAll } from 'vitest';
import { render, screen, fireEvent, cleanup } from '@testing-library/react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { TextPart } from './TextPart';
import { ExecuteCodePart } from './ExecuteCodePart';
import { StepStartPart } from './StepStartPart';
import { LookupSkillPart } from './LookupSkillPart';
import { DynamicToolPart } from './DynamicToolPart';
import { ApprovalRequestedPart } from './ApprovalRequestedPart';
import { LanguageProvider } from '../../i18n/index';
import { translationService } from '../../i18n/index';

beforeAll(async () => {
  await translationService.setLocale('en');
});

afterEach(() => cleanup());

function renderWithFluent(ui: React.ReactElement) {
  return render(
    <LanguageProvider initialLocale="en">
      <FluentProvider theme={webLightTheme}>{ui}</FluentProvider>
    </LanguageProvider>,
  );
}

describe('TextPart', () => {
  it('renders text', () => {
    render(<TextPart part={{ text: 'hello' }} />);
    expect(screen.getByText('hello')).toBeInTheDocument();
  });

  it('renders safe Markdown and mixed Arabic/Latin text', () => {
    const { container } = render(
      <TextPart part={{ text: '**تم** تحديث `Q3` <script>alert(1)</script>' }} />,
    );
    expect(screen.getByText('تم')).toHaveTextContent('تم');
    expect(screen.getByText('Q3')).toHaveAttribute('dir', 'ltr');
    expect(container.querySelector('script')).toBeNull();
  });
});

describe('StepStartPart', () => {
  it('renders an hr', () => {
    const { container } = render(<StepStartPart />);
    expect(container.querySelector('hr')).not.toBeNull();
  });
});

describe('ExecuteCodePart', () => {
  it('shows Approve when state is input-available', () => {
    const onApprove = vi.fn();
    renderWithFluent(
      <ExecuteCodePart
        part={{ state: 'input-available', toolCallId: 'tc', input: { code: 'await 1' } }}
        onApprove={onApprove}
        onReject={() => {}}
        highlight={(s) => s}
      />,
    );
    fireEvent.click(screen.getByRole('button', { name: /Approve & Run/i }));
    expect(onApprove).toHaveBeenCalledWith('tc', 'await 1');
  });

  it('shows output-error message in error state', () => {
    renderWithFluent(
      <ExecuteCodePart
        part={{ state: 'output-error', toolCallId: 'tc', errorText: 'kaboom' }}
        onApprove={() => {}}
        onReject={() => {}}
        highlight={(s) => s}
      />,
    );
    expect(screen.getByText('kaboom')).toBeInTheDocument();
  });

  it('shows a user-facing change summary before code', () => {
    renderWithFluent(
      <ExecuteCodePart
        part={{
          state: 'input-available',
          toolCallId: 'tc',
          input: { summary: 'تنسيق العناوين', code: 'await context.sync()' },
        }}
        onApprove={() => {}}
        onReject={() => {}}
        highlight={(s) => s}
      />,
    );
    expect(screen.getByText('Planned change')).toBeInTheDocument();
    expect(screen.getByText('تنسيق العناوين')).toHaveAttribute('dir', 'auto');
  });
});

describe('LookupSkillPart', () => {
  it('shows looking-up label when in flight', () => {
    render(<LookupSkillPart part={{ state: 'input-available', input: { name: 'tables' } }} />);
    expect(screen.getByText(/Looking up: tables/)).toBeInTheDocument();
  });

  it('shows looked-up label when complete', () => {
    render(
      <LookupSkillPart
        part={{ state: 'output-available', input: { name: 'ranges' }, output: { body: '...' } }}
      />,
    );
    expect(screen.getByText(/Looked up: ranges/)).toBeInTheDocument();
  });
});

describe('DynamicToolPart', () => {
  it('renders tool name and a friendly completed state', () => {
    render(<DynamicToolPart part={{ toolName: 'mcp_x/list', state: 'output-available', input: {} }} />);
    expect(screen.getByText('mcp_x/list')).toBeInTheDocument();
    expect(screen.getByText('Done')).toBeInTheDocument();
  });
});

describe('ApprovalRequestedPart', () => {
  it('renders completed tool output as a regular tool part', () => {
    const { container } = render(
      <ApprovalRequestedPart
        part={{ type: 'tool-x', state: 'output-available' }}
        onResponse={() => {}}
      />,
    );
    expect(container.firstChild).not.toBeNull();
    expect(screen.getByText('Done')).toBeInTheDocument();
  });

  it('calls onResponse with approved=true on Approve click', () => {
    const onResponse = vi.fn();
    render(
      <ApprovalRequestedPart
        part={{
          type: 'tool-foo',
          state: 'approval-requested',
          approval: { id: 'a1' },
          input: { x: 1 },
        }}
        onResponse={onResponse}
      />,
    );
    fireEvent.click(screen.getByRole('button', { name: 'Approve' }));
    expect(onResponse).toHaveBeenCalledWith('a1', true);
  });

  it('calls onResponse with approved=false on Deny click', () => {
    const onResponse = vi.fn();
    render(
      <ApprovalRequestedPart
        part={{
          type: 'tool-foo',
          state: 'approval-requested',
          approval: { id: 'a1' },
          input: {},
        }}
        onResponse={onResponse}
      />,
    );
    fireEvent.click(screen.getByRole('button', { name: 'Deny' }));
    expect(onResponse).toHaveBeenCalledWith('a1', false);
  });
});
