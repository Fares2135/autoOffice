import React, { useState, useRef, useEffect, useLayoutEffect } from 'react';
import {
  makeStyles,
  tokens,
  Badge,
  Textarea,
  Button,
  Text,
  Tooltip,
} from '@fluentui/react-components';
import {
  Send24Regular,
  Settings24Regular,
  History24Regular,
  Add24Regular,
} from '@fluentui/react-icons';
import type { HostContext } from '../host/context.ts';
import { useTranslation } from '../i18n/index.ts';
import { MessageBubble, type UIMessageLike } from './MessageBubble.tsx';
import {
  UsageCostSchema,
  sumUsageCosts,
  type ProviderConfig,
} from '@autooffice/shared';
import { CostBadge } from './CostBadge.tsx';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    overflow: 'hidden',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '8px 12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
  },
  brand: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  logo: {
    width: '24px',
    height: '24px',
    flexShrink: 0,
  },
  title: {
    fontWeight: 600,
    fontSize: '16px',
  },
  messageList: {
    flex: 1,
    overflowY: 'auto',
    padding: '8px 0',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  empty: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'stretch',
    justifyContent: 'center',
    flex: 1,
    gap: '8px',
    color: tokens.colorNeutralForeground3,
    padding: '24px',
    textAlign: 'center',
  },
  welcomeCard: {
    position: 'relative',
    overflow: 'hidden',
    padding: '20px',
    borderRadius: '16px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundImage: `radial-gradient(circle at top right, ${tokens.colorBrandBackground2} 0, transparent 48%)`,
    backgroundColor: tokens.colorNeutralBackground2,
    boxShadow: tokens.shadow4,
  },
  eyebrow: {
    display: 'block',
    color: tokens.colorBrandForeground1,
    fontSize: '11px',
    fontWeight: 700,
    letterSpacing: '0.12em',
    marginBottom: '8px',
    textTransform: 'uppercase',
  },
  stepRail: {
    display: 'grid',
    gap: '8px',
    marginTop: '18px',
    textAlign: 'start',
  },
  step: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    color: tokens.colorNeutralForeground2,
  },
  stepNumber: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '24px',
    height: '24px',
    borderRadius: '50%',
    color: tokens.colorBrandForeground1,
    backgroundColor: tokens.colorBrandBackground2,
    fontSize: '12px',
    fontWeight: 700,
    flexShrink: 0,
  },
  promptHeading: {
    marginTop: '20px',
    marginBottom: '4px',
    textAlign: 'start',
    color: tokens.colorNeutralForeground2,
  },
  promptGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(190px, 1fr))',
    gap: '8px',
  },
  promptCard: {
    width: '100%',
    minHeight: '84px',
    padding: '12px',
    textAlign: 'start',
    justifyContent: 'flex-start',
    alignItems: 'flex-start',
    borderRadius: '12px',
    whiteSpace: 'normal',
    transitionProperty: 'transform, box-shadow, border-color',
    transitionDuration: '160ms',
    ':hover': {
      transform: 'translateY(-2px)',
      boxShadow: tokens.shadow8,
      borderTopColor: tokens.colorBrandStroke1,
      borderRightColor: tokens.colorBrandStroke1,
      borderBottomColor: tokens.colorBrandStroke1,
      borderLeftColor: tokens.colorBrandStroke1,
    },
    ':focus-visible': {
      outline: `2px solid ${tokens.colorStrokeFocus2}`,
      outlineOffset: '2px',
    },
  },
  promptCopy: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    alignItems: 'flex-start',
  },
  inputArea: {
    display: 'flex',
    alignItems: 'flex-end',
    gap: '8px',
    padding: '8px 12px',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
  },
  banner: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 12px',
    backgroundColor: tokens.colorPaletteYellowBackground1,
    borderBottom: `1px solid ${tokens.colorPaletteYellowBorder1}`,
    color: tokens.colorPaletteDarkOrangeForeground1,
    flexShrink: 0,
  },
  bannerError: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    borderBottomColor: tokens.colorPaletteRedBorder1,
    color: tokens.colorPaletteRedForeground1,
  },
  bannerText: {
    flex: 1,
    minWidth: 0,
  },
  input: {
    flex: 1,
  },
  textarea: {
    width: '100%',
    minHeight: '32px',
    maxHeight: '200px',
    resize: 'none',
    overflowY: 'auto',
  },
});

export interface ChatPanelProps {
  host: HostContext;
  messages: UIMessageLike[];
  status: 'submitted' | 'streaming' | 'ready' | 'error' | string;
  /** Last error from the chat stream, if any. Surfaced as an inline banner. */
  chatError?: string | null;
  /** True when no provider is configured/selected; disables send and shows banner. */
  noProvider?: boolean;
  providers?: ProviderConfig[];
  autoApprove?: boolean;
  runInIframe?: (code: string) => Promise<unknown>;
  onCliResolved?: (resultText: string) => void;
  onSubmit: (text: string) => void;
  onApproveCode: (toolCallId: string, code: string) => Promise<void> | void;
  onRejectCode: (toolCallId: string) => void;
  onApprovalResponse: (id: string, approved: boolean) => void;
  highlightCode: (code: string) => React.ReactNode;
  onOpenSettings?: () => void;
  onOpenHistory?: () => void;
  onNewChat?: () => void;
}

function isProviderError(msg: string): boolean {
  const m = msg.toLowerCase();
  return (
    m.includes('api key') || m.includes('api_key') ||
    m.includes('unauthorized') || m.includes('authentication') ||
    m.includes('401') || m.includes('403') ||
    m.includes('no provider') || m.includes('provider not found') ||
    m.includes('no model') || m.includes('invalid key') ||
    m.includes('quota') || m.includes('billing') ||
    m.includes('permission') || m.includes('credentials')
  );
}

export function ChatPanel({
  host,
  messages,
  status,
  chatError,
  noProvider,
  providers = [],
  autoApprove = false,
  runInIframe,
  onCliResolved,
  onSubmit,
  onApproveCode,
  onRejectCode,
  onApprovalResponse,
  highlightCode,
  onOpenSettings,
  onOpenHistory,
  onNewChat,
}: ChatPanelProps) {
  const styles = useStyles();
  const { t } = useTranslation();
  const isLoading = status === 'submitted' || status === 'streaming';
  const hostDisplay = t(
    host.kind === 'word'
      ? 'chat.hostWord'
      : host.kind === 'excel'
        ? 'chat.hostExcel'
        : 'chat.hostPowerpoint',
  );
  const hostNoun = t(
    host.kind === 'word'
      ? 'chat.hostNounWord'
      : host.kind === 'excel'
        ? 'chat.hostNounExcel'
        : 'chat.hostNounPowerpoint',
  );
  const [inputText, setInputText] = useState('');
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const prompts = host.kind === 'word'
    ? [
        [t('chat.promptWordFormat'), t('chat.promptWordFormatBody')],
        [t('chat.promptWordSummarize'), t('chat.promptWordSummarizeBody')],
      ]
    : host.kind === 'excel'
      ? [
          [t('chat.promptExcelClean'), t('chat.promptExcelCleanBody')],
          [t('chat.promptExcelAnalyze'), t('chat.promptExcelAnalyzeBody')],
        ]
      : [
          [t('chat.promptPowerpointPolish'), t('chat.promptPowerpointPolishBody')],
        [t('chat.promptPowerpointOutline'), t('chat.promptPowerpointOutlineBody')],
      ];
  const usageCost = sumUsageCosts(
    messages.flatMap((message) => {
      const parsed = UsageCostSchema.safeParse(message.metadata?.usageCost);
      return parsed.success ? [parsed.data] : [];
    }),
  );

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  useLayoutEffect(() => {
    const el = textareaRef.current;
    if (!el) return;
    el.style.height = 'auto';
    el.style.height = `${Math.min(el.scrollHeight, 200)}px`;
  }, [inputText]);

  const handleSubmit = () => {
    if (!inputText.trim() || isLoading || noProvider) return;
    onSubmit(inputText);
    setInputText('');
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <div className={styles.brand}>
          <img
            src={`${import.meta.env.BASE_URL}assets/icon-64.png`}
            alt=""
            className={styles.logo}
          />
          <Text className={styles.title}>AutoOffice</Text>
          <Badge
            appearance="outline"
            size="small"
            color={host.kind === 'excel' ? 'success' : host.kind === 'powerpoint' ? 'danger' : 'brand'}
          >
            {host.displayName}
          </Badge>
          <CostBadge cost={usageCost} />
        </div>
        <div style={{ display: 'flex', gap: '4px' }}>
          {onOpenHistory && (
            <Tooltip content={t('chat.historyTooltip')} relationship="label">
              <Button
                appearance="subtle"
                icon={<History24Regular />}
                onClick={onOpenHistory}
                disabled={isLoading}
              />
            </Tooltip>
          )}
          {onNewChat && (
            <Tooltip content={t('chat.newChatTooltip')} relationship="label">
              <Button
                appearance="subtle"
                icon={<Add24Regular />}
                onClick={onNewChat}
                disabled={isLoading}
              />
            </Tooltip>
          )}
          {onOpenSettings && (
            <Tooltip content={t('chat.settingsTooltip')} relationship="label">
              <Button appearance="subtle" icon={<Settings24Regular />} onClick={onOpenSettings} />
            </Tooltip>
          )}
        </div>
      </div>

      {chatError && !noProvider && (
        <div
          className={`${styles.banner} ${styles.bannerError}`}
          role="alert"
          aria-label={t('chat.chatErrorAria')}
        >
          <Text className={styles.bannerText} size={200}>
            {chatError}
          </Text>
          {onOpenSettings && isProviderError(chatError) && (
            <Button appearance="subtle" size="small" onClick={onOpenSettings}>
              {t('chat.settingsTooltip')}
            </Button>
          )}
        </div>
      )}

      <div className={styles.messageList}>
        {messages.length === 0 ? (
          <div className={styles.empty}>
            {noProvider ? (
              <section className={styles.welcomeCard} aria-labelledby="onboarding-title">
                <Text className={styles.eyebrow}>{t('chat.onboardingEyebrow')}</Text>
                <Text id="onboarding-title" as="h1" size={500} weight="semibold" block>
                  {t('chat.onboardingTitle')}
                </Text>
                <Text size={200} block>
                  {t('chat.onboardingBody')}
                </Text>
                <div className={styles.stepRail}>
                  {[
                    t('chat.onboardingStepProvider'),
                    t('chat.onboardingStepKey'),
                    t('chat.onboardingStepModel'),
                  ].map((label, index) => (
                    <div className={styles.step} key={label}>
                      <span className={styles.stepNumber}>{index + 1}</span>
                      <Text size={200}>{label}</Text>
                    </div>
                  ))}
                </div>
                {onOpenSettings && (
                  <Button
                    appearance="primary"
                    onClick={onOpenSettings}
                    style={{ marginTop: '18px' }}
                  >
                    {t('chat.openSettings')}
                  </Button>
                )}
              </section>
            ) : (
              <>
                <section className={styles.welcomeCard}>
                  <Text className={styles.eyebrow}>{hostDisplay}</Text>
                  <Text as="h1" size={500} weight="semibold" block>
                    {t('chat.welcomeTitle')}
                  </Text>
                  <Text size={200} block>
                    {t('chat.welcomeMessage', { host: hostDisplay, noun: hostNoun })}
                  </Text>
                </section>
                <Text className={styles.promptHeading} size={200} weight="semibold">
                  {t('chat.promptHeading')}
                </Text>
                <div className={styles.promptGrid}>
                  {prompts.map(([title, body]) => (
                    <Button
                      key={title}
                      appearance="outline"
                      className={styles.promptCard}
                      onClick={() => {
                        setInputText(title);
                        requestAnimationFrame(() => textareaRef.current?.focus());
                      }}
                    >
                      <span className={styles.promptCopy}>
                        <Text weight="semibold" dir="auto">{title}</Text>
                        <Text size={200} dir="auto">{body}</Text>
                      </span>
                    </Button>
                  ))}
                </div>
              </>
            )}
          </div>
        ) : (
          (() => {
            // Detect synthetic tool_results messages (any text part starting with <tool_results>).
            const isSynthetic = (m: UIMessageLike) =>
              m.role === 'user' &&
              m.parts.some(p => typeof (p as any).text === 'string' && (p as any).text.startsWith('<tool_results>'));

            // Build set of assistant message IDs that already have a tool_results reply.
            const resolvedIds = new Set<string>();
            for (let i = 0; i < messages.length - 1; i++) {
              const cur = messages[i]!;
              const next = messages[i + 1]!;
              if (cur.role === 'assistant' && isSynthetic(next) && cur.id) {
                resolvedIds.add(cur.id);
              }
            }

            return messages
              .filter(m => !isSynthetic(m))
              .map((msg, i, arr) => (
                <MessageBubble
                  key={msg.id ?? i}
                  message={msg}
                  onApproveCode={onApproveCode}
                  onRejectCode={onRejectCode}
                  onApprovalResponse={onApprovalResponse}
                  highlightCode={highlightCode}
                  streaming={isLoading && i === arr.length - 1 && msg.role === 'assistant'}
                  providers={providers}
                  onCliResolved={onCliResolved}
                  autoApprove={autoApprove}
                  runInIframe={runInIframe}
                  cliHistorical={resolvedIds.has(msg.id ?? '')}
                />
              ));
          })()
        )}
        <div ref={messagesEndRef} />
      </div>

      <div className={styles.inputArea}>
        <Textarea
          className={styles.input}
          textarea={{ ref: textareaRef, className: styles.textarea }}
          dir="auto"
          placeholder={t('chat.inputPlaceholder', { noun: hostNoun })}
          value={inputText}
          onChange={(_, data) => setInputText(data.value)}
          onKeyDown={handleKeyDown}
          disabled={isLoading}
          rows={1}
        />
        <Button
          appearance="primary"
          icon={<Send24Regular />}
          aria-label={t('chat.sendButton')}
          onClick={handleSubmit}
          disabled={!inputText.trim() || isLoading || !!noProvider}
        />
      </div>
    </div>
  );
}
