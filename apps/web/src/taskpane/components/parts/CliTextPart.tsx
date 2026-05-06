import React, { useEffect, useRef, useState } from 'react';
import { makeStyles, tokens, Button, Badge, Text, Tooltip, Spinner } from '@fluentui/react-components';
import { Play24Regular, DismissCircle24Regular } from '@fluentui/react-icons';
import { apiGet } from '../../api.ts';
import { parseCliSegments, buildToolResults, hasCliTools } from '../../chat/cli-tool-parser.ts';
import type { CliSegment } from '../../chat/cli-tool-parser.ts';

export { hasCliTools };

const useStyles = makeStyles({
  codeBlock: {
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: '8px',
    overflow: 'hidden',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    maxWidth: '100%',
    marginTop: '4px',
    marginBottom: '4px',
  },
  codeHeader: {
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
  details: { borderTop: `1px solid ${tokens.colorNeutralStroke1}` },
  detailsError: { backgroundColor: tokens.colorPaletteRedBackground1 },
  summary: {
    padding: '6px 12px',
    cursor: 'pointer',
    fontSize: '12px',
    fontWeight: 600,
    color: tokens.colorNeutralForeground2,
    userSelect: 'none',
  },
  summaryError: { color: tokens.colorPaletteRedForeground1 },
  resultBody: {
    padding: '8px 12px 12px',
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
  },
  skillSpan: { opacity: 0.7, fontSize: '12px', direction: 'ltr', display: 'block' },
});

type BlockStatus = 'pending' | 'running' | 'done' | 'error';
type BlockState = { status: BlockStatus; output?: unknown; error?: string };

type Props = {
  text: string;
  /** True when this message already has a tool_results reply in history — show read-only, no auto-run. */
  isHistorical: boolean;
  autoApprove: boolean;
  runInIframe: (code: string) => Promise<unknown>;
  highlightCode: (code: string) => React.ReactNode;
  onAllResolved: (resultText: string) => void;
};

export function CliTextPart({ text, isHistorical, autoApprove, runInIframe, highlightCode, onAllResolved }: Props) {
  const styles = useStyles();
  const segments = parseCliSegments(text);
  const toolSegments = segments.filter(s => s.type !== 'text') as Exclude<CliSegment, { type: 'text' }>[];

  const [blockStates, setBlockStates] = useState<Map<string, BlockState>>(() => {
    const m = new Map<string, BlockState>();
    for (const seg of toolSegments) {
      const key = seg.type === 'lookup_skill' ? `lookup_skill:${seg.name}` : `execute_code:${seg.index}`;
      m.set(key, { status: 'pending' });
    }
    return m;
  });

  const resolvedRef = useRef(false);
  // Guard against React StrictMode double-invoke on mount effects.
  const startedRef = useRef(false);

  // On mount (once): kick off all auto-resolution unless this is a historical message.
  useEffect(() => {
    if (isHistorical || startedRef.current) return;
    startedRef.current = true;

    for (const seg of toolSegments) {
      if (seg.type === 'lookup_skill') {
        const key = `lookup_skill:${seg.name}`;
        setBlockStates(prev => {
          if (prev.get(key)?.status !== 'pending') return prev;
          const next = new Map(prev);
          next.set(key, { status: 'running' });
          return next;
        });
        apiGet<{ name: string; body: string } | { error: string }>(`/api/skills/${seg.name}`)
          .then(res => {
            setBlockStates(prev => {
              const next = new Map(prev);
              next.set(key, 'error' in res
                ? { status: 'error', error: res.error }
                : { status: 'done', output: res.body });
              return next;
            });
          })
          .catch(err => {
            setBlockStates(prev => {
              const next = new Map(prev);
              next.set(key, { status: 'error', error: (err as Error).message });
              return next;
            });
          });
      } else if (seg.type === 'execute_code' && autoApprove) {
        runBlock(`execute_code:${seg.index}`, seg.code);
      }
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // When all tool blocks reach a terminal state, fire onAllResolved exactly once.
  useEffect(() => {
    if (isHistorical || toolSegments.length === 0 || resolvedRef.current) return;
    const allDone = toolSegments.every(seg => {
      const key = seg.type === 'lookup_skill' ? `lookup_skill:${seg.name}` : `execute_code:${seg.index}`;
      const s = blockStates.get(key)?.status;
      return s === 'done' || s === 'error';
    });
    if (!allDone) return;
    resolvedRef.current = true;
    const results = new Map<string, { success: boolean; output?: unknown; error?: string }>();
    for (const [key, state] of blockStates) {
      if (state.status === 'done') results.set(key, { success: true, output: state.output });
      else if (state.status === 'error') results.set(key, { success: false, error: state.error });
    }
    onAllResolved(buildToolResults(segments, results));
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [blockStates]);

  function runBlock(key: string, code: string) {
    setBlockStates(prev => {
      if (prev.get(key)?.status !== 'pending') return prev;
      const next = new Map(prev);
      next.set(key, { status: 'running' });
      return next;
    });
    runInIframe(code)
      .then(output => setBlockStates(prev => { const n = new Map(prev); n.set(key, { status: 'done', output }); return n; }))
      .catch(err => setBlockStates(prev => { const n = new Map(prev); n.set(key, { status: 'error', error: (err as Error).message }); return n; }));
  }

  function rejectBlock(key: string) {
    setBlockStates(prev => { const n = new Map(prev); n.set(key, { status: 'error', error: 'User rejected' }); return n; });
  }

  return (
    <>
      {segments.map((seg, i) => {
        if (seg.type === 'text') {
          return <span key={i}>{seg.content}</span>;
        }

        if (seg.type === 'lookup_skill') {
          const state = blockStates.get(`lookup_skill:${seg.name}`);
          return (
            <span key={i} className={styles.skillSpan}>
              {state?.status === 'done' ? `Looked up: ${seg.name}` : `Looking up: ${seg.name}`}
              {state?.status === 'running' && <Spinner size="extra-tiny" style={{ marginLeft: 4 }} />}
            </span>
          );
        }

        // execute_code
        const key = `execute_code:${seg.index}`;
        const state = blockStates.get(key) ?? { status: 'pending' };
        const showApprove = !isHistorical && state.status === 'pending';
        const statusLabel =
          state.status === 'running' ? 'Running…' :
          state.status === 'done' ? 'Success' :
          state.status === 'error' ? 'Error' :
          isHistorical ? 'Ran' : 'Awaiting Approval';
        const badgeColor: 'informative' | 'success' | 'danger' =
          state.status === 'done' ? 'success' :
          state.status === 'error' ? 'danger' : 'informative';
        const showResult =
          (state.status === 'done' && state.output !== undefined) ||
          (state.status === 'error' && !!state.error);
        const resultText = state.status === 'error'
          ? state.error ?? ''
          : state.output !== undefined
            ? typeof state.output === 'string' ? state.output : JSON.stringify(state.output, null, 2)
            : '';

        return (
          <div key={i} className={styles.codeBlock} dir="ltr">
            <div className={styles.codeHeader}>
              <Text size={200} weight="semibold">office.js</Text>
              <Badge appearance="filled" color={badgeColor}>{statusLabel}</Badge>
            </div>
            <div className={styles.codeArea} style={{ direction: 'ltr', textAlign: 'left' }}>{highlightCode(seg.code)}</div>
            {showApprove && (
              <div className={styles.actions}>
                <Tooltip content="Approve & Run" relationship="label" withArrow>
                  <Button
                    className={styles.approveBtn}
                    icon={<Play24Regular />}
                    size="small"
                    aria-label="Approve & Run"
                    onClick={() => runBlock(key, seg.code)}
                  />
                </Tooltip>
                <Tooltip content="Reject" relationship="label" withArrow>
                  <Button
                    appearance="subtle"
                    icon={<DismissCircle24Regular />}
                    size="small"
                    aria-label="Reject"
                    onClick={() => rejectBlock(key)}
                  />
                </Tooltip>
              </div>
            )}
            {showResult && (
              <details
                className={`${styles.details} ${state.status === 'error' ? styles.detailsError : ''}`}
                open
              >
                <summary className={`${styles.summary} ${state.status === 'error' ? styles.summaryError : ''}`} style={{ direction: 'ltr', textAlign: 'left' }}>
                  {state.status === 'error' ? 'Error details' : 'Result'}
                </summary>
                <div className={`${styles.resultBody} ${state.status === 'error' ? styles.resultBodyError : ''}`} style={{ direction: 'ltr', textAlign: 'left' }}>
                  {resultText}
                </div>
              </details>
            )}
          </div>
        );
      })}
    </>
  );
}
