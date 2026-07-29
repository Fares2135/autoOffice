import React from 'react';
import { makeStyles, tokens, Button, Badge, Text, Tooltip } from '@fluentui/react-components';
import { DismissCircle24Regular, Play24Regular } from '@fluentui/react-icons';
import { useTranslation } from '../../i18n/index.ts';

const useStyles = makeStyles({
  container: {
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: '8px',
    overflow: 'hidden',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    maxWidth: '100%',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '6px 12px',
    backgroundColor: tokens.colorNeutralBackground4,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  codeArea: {
    overflowX: 'auto',
    overflowY: 'auto',
    maxHeight: '300px',
    fontFamily: 'Consolas, "Courier New", monospace',
    fontSize: '12px',
    lineHeight: '1.5',
    color: tokens.colorNeutralForeground1,
    direction: 'ltr',
    textAlign: 'left',
    '& pre': {
      margin: 0,
      padding: '12px',
      display: 'block',
      minWidth: 'max-content',
      whiteSpace: 'pre',
      boxSizing: 'border-box',
      direction: 'ltr',
      textAlign: 'left',
    },
    '& code': {
      whiteSpace: 'pre',
      direction: 'ltr',
      textAlign: 'left',
    },
  },
  actions: {
    display: 'flex',
    gap: '8px',
    padding: '8px 12px',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  details: {
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  detailsError: {
    backgroundColor: tokens.colorPaletteRedBackground1,
  },
  summary: {
    padding: '6px 12px',
    cursor: 'pointer',
    fontSize: '12px',
    fontWeight: 600,
    color: tokens.colorNeutralForeground2,
    userSelect: 'none',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground4Hover,
    },
  },
  summaryError: {
    color: tokens.colorPaletteRedForeground1,
    '&:hover': {
      backgroundColor: tokens.colorPaletteRedBackground2,
    },
  },
  resultBody: {
    padding: '8px 12px 12px 12px',
    fontSize: '12px',
    fontFamily: 'Consolas, "Courier New", monospace',
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
    maxHeight: '300px',
    overflow: 'auto',
    color: tokens.colorNeutralForeground1,
    direction: 'ltr',
    textAlign: 'left',
  },
  resultBodyError: {
    color: tokens.colorPaletteRedForeground1,
  },
  approveBtn: {
    backgroundColor: tokens.colorPaletteGreenBackground3,
    borderTopColor: tokens.colorPaletteGreenBorderActive,
    borderRightColor: tokens.colorPaletteGreenBorderActive,
    borderBottomColor: tokens.colorPaletteGreenBorderActive,
    borderLeftColor: tokens.colorPaletteGreenBorderActive,
    color: tokens.colorNeutralForegroundOnBrand,
    '&:hover': {
      backgroundColor: tokens.colorPaletteGreenForeground1,
      borderTopColor: tokens.colorPaletteGreenForeground1,
      borderRightColor: tokens.colorPaletteGreenForeground1,
      borderBottomColor: tokens.colorPaletteGreenForeground1,
      borderLeftColor: tokens.colorPaletteGreenForeground1,
      color: tokens.colorNeutralForegroundOnBrand,
    },
    '&:hover:active': {
      backgroundColor: tokens.colorPaletteGreenForeground3,
      borderTopColor: tokens.colorPaletteGreenForeground3,
      borderRightColor: tokens.colorPaletteGreenForeground3,
      borderBottomColor: tokens.colorPaletteGreenForeground3,
      borderLeftColor: tokens.colorPaletteGreenForeground3,
      color: tokens.colorNeutralForegroundOnBrand,
    },
  },
  changeSummary: {
    padding: '10px 12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorNeutralForeground1,
    textAlign: 'start',
  },
});

type CodeStatus = 'streaming' | 'pending' | 'success' | 'error';

const STATUS_COLORS: Record<CodeStatus, 'informative' | 'success' | 'danger'> = {
  streaming: 'informative',
  pending: 'informative',
  success: 'success',
  error: 'danger',
};

function statusFromState(state: string): CodeStatus {
  switch (state) {
    case 'input-streaming':
      return 'streaming';
    case 'output-error':
      return 'error';
    case 'output-available':
      return 'success';
    case 'input-available':
    default:
      return 'pending';
  }
}

type Props = {
  part: {
    state: string;
    toolCallId: string;
    input?: { code?: string; summary?: string };
    output?: unknown;
    errorText?: string;
  };
  onApprove: (toolCallId: string, code: string) => void;
  onReject: (toolCallId: string) => void;
  highlight: (code: string) => React.ReactNode;
};

export function ExecuteCodePart({ part, onApprove, onReject, highlight }: Props) {
  const styles = useStyles();
  const { t } = useTranslation();
  const code = part.input?.code ?? '';
  const changeSummary = part.input?.summary?.trim();
  const status = statusFromState(part.state);
  const isError = status === 'error';
  const showResult =
    (status === 'success' && part.output !== undefined) ||
    (status === 'error' && !!part.errorText);
  const resultText = isError
    ? part.errorText ?? ''
    : part.output !== undefined
    ? typeof part.output === 'string'
      ? part.output
      : JSON.stringify(part.output, null, 2)
    : '';

  return (
    <div className={styles.container} dir="ltr">
      <div className={styles.header}>
        <Text size={200} weight="semibold">office.js</Text>
        <Badge appearance="filled" color={STATUS_COLORS[status]}>
          {t(
            status === 'streaming'
              ? 'code.statusStreaming'
              : status === 'pending'
                ? 'code.statusPending'
                : status === 'success'
                  ? 'code.statusSuccess'
                  : 'code.statusError',
          )}
        </Badge>
      </div>

      {changeSummary && (
        <div className={styles.changeSummary} dir="auto">
          <Text size={100} weight="semibold" block>
            {t('code.changeSummary')}
          </Text>
          <Text size={200} dir="auto">{changeSummary}</Text>
        </div>
      )}

      <div className={styles.codeArea} style={{ direction: 'ltr', textAlign: 'left' }}>{highlight(code)}</div>

      {status === 'pending' && (
        <div className={styles.actions}>
          <Tooltip content={t('code.approveButton')} relationship="label" withArrow>
            <Button
              className={styles.approveBtn}
              icon={<Play24Regular />}
              size="small"
              aria-label={t('code.approveButton')}
              onClick={() => onApprove(part.toolCallId, code)}
            />
          </Tooltip>
          <Tooltip content={t('code.rejectButton')} relationship="label" withArrow>
            <Button
              appearance="subtle"
              icon={<DismissCircle24Regular />}
              size="small"
              aria-label={t('code.rejectButton')}
              onClick={() => onReject(part.toolCallId)}
            />
          </Tooltip>
        </div>
      )}

      {showResult && (
        <details
          className={`${styles.details} ${isError ? styles.detailsError : ''}`}
          open={isError}
        >
          <summary className={`${styles.summary} ${isError ? styles.summaryError : ''}`} style={{ direction: 'ltr', textAlign: 'left' }}>
            {isError ? t('code.errorDetails') : t('code.result')}
          </summary>
          <div
            className={`${styles.resultBody} ${isError ? styles.resultBodyError : ''}`}
            style={{ direction: 'ltr', textAlign: 'left' }}
          >
            {resultText}
          </div>
        </details>
      )}
    </div>
  );
}
