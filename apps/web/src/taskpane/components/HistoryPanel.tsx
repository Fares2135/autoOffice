import React, { useEffect, useState, useCallback } from 'react';
import {
  makeStyles,
  tokens,
  Button,
  Text,
  Badge,
  Input,
  Switch,
  Spinner,
} from '@fluentui/react-components';
import {
  Dismiss24Regular,
  Edit20Regular,
  Delete20Regular,
  Checkmark20Regular,
  DismissCircle20Regular,
} from '@fluentui/react-icons';
import { apiGet, apiSend } from '../api.ts';
import type { Conversation, Host } from '@autooffice/shared';
import { useFormatters, useTranslation } from '../i18n/index.ts';
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
    gap: '8px',
    padding: '8px 12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
  },
  filters: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
  },
  list: {
    flex: 1,
    overflowY: 'auto',
  },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    cursor: 'pointer',
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  rowActive: {
    backgroundColor: tokens.colorNeutralBackground1Selected,
  },
  rowMain: {
    flex: 1,
    minWidth: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  rowTitle: {
    fontWeight: 500,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  rowMeta: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    color: tokens.colorNeutralForeground3,
    fontSize: '12px',
  },
  rowActions: {
    display: 'flex',
    gap: '2px',
  },
  empty: {
    padding: '24px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
  },
  loading: {
    display: 'flex',
    justifyContent: 'center',
    padding: '24px',
  },
  error: {
    padding: '12px 16px',
    color: tokens.colorPaletteRedForeground1,
  },
});

export interface HistoryPanelProps {
  /** The current host's kind — used to filter by default. */
  currentHost: Host;
  /** The currently-active conversation id, if any. */
  activeConversationId: string | null;
  /** Fired when the user picks a conversation. */
  onSelectConversation: (id: string) => void;
  /** Fired when the user closes the panel. */
  onClose: () => void;
  /**
   * Optional override for the load function — used by tests. In production this
   * defaults to apiGet('/api/conversations').
   */
  loadConversations?: () => Promise<Conversation[]>;
}

function relativeTimeParts(ts: number): [number, Intl.RelativeTimeFormatUnit] | null {
  const minutes = Math.round((ts - Date.now()) / 60_000);
  const absMin = Math.abs(minutes);
  if (absMin < 1) return null;
  if (absMin < 60) return [minutes, 'minute'];
  const hours = Math.round(minutes / 60);
  if (Math.abs(hours) < 24) return [hours, 'hour'];
  return [Math.round(hours / 24), 'day'];
}

export function HistoryPanel({
  currentHost,
  activeConversationId,
  onSelectConversation,
  onClose,
  loadConversations,
}: HistoryPanelProps) {
  const styles = useStyles();
  const { t } = useTranslation();
  const { formatRelativeTime } = useFormatters();
  const [conversations, setConversations] = useState<Conversation[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [showAll, setShowAll] = useState(false);
  const [renamingId, setRenamingId] = useState<string | null>(null);
  const [renameDraft, setRenameDraft] = useState('');

  const reload = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const list = loadConversations
        ? await loadConversations()
        : await apiGet<Conversation[]>('/api/conversations');
      setConversations(list);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLoading(false);
    }
  }, [loadConversations]);

  useEffect(() => {
    void reload();
  }, [reload]);

  const filtered = (showAll ? conversations : conversations.filter((c) => c.host === currentHost))
    .slice()
    .sort((a, b) => b.updatedAt - a.updatedAt);
  const hostLabel = (host: Host) => t(
    host === 'word'
      ? 'history.filterWord'
      : host === 'excel'
        ? 'history.filterExcel'
        : 'history.filterPowerpoint',
  );
  const relativeTime = (timestamp: number) => {
    const parts = relativeTimeParts(timestamp);
    return parts ? formatRelativeTime(parts[0], parts[1]) : t('history.relativeNow');
  };

  const startRename = (c: Conversation) => {
    setRenamingId(c.id);
    setRenameDraft(c.title ?? '');
  };

  const commitRename = async () => {
    if (renamingId && renameDraft.trim()) {
      try {
        await apiSend(`/api/conversations/${renamingId}`, { title: renameDraft.trim() }, 'PATCH');
        await reload();
      } catch (e) {
        setError((e as Error).message);
      }
    }
    setRenamingId(null);
    setRenameDraft('');
  };

  const cancelRename = () => {
    setRenamingId(null);
    setRenameDraft('');
  };

  const requestDelete = async (id: string) => {
    try {
      await apiSend(`/api/conversations/${id}`, null, 'DELETE');
      await reload();
    } catch (e) {
      setError((e as Error).message);
    }
  };

  return (
    <div className={styles.container} role="dialog" aria-label={t('history.dialogAria')}>
      <div className={styles.header}>
        <Button
          appearance="subtle"
          icon={<Dismiss24Regular />}
          onClick={onClose}
          aria-label={t('history.closeAria')}
        />
        <Text weight="semibold">{t('history.title')}</Text>
      </div>

      <div className={styles.filters}>
        <Switch
          checked={showAll}
          onChange={(_, d) => setShowAll(d.checked)}
          label={t('history.showAllHosts')}
        />
        {!showAll && (
          <Badge appearance="outline" size="small">
            {hostLabel(currentHost)}
          </Badge>
        )}
      </div>

      {error && <div className={styles.error}>{error}</div>}

      <div className={styles.list}>
        {loading ? (
          <div className={styles.loading}>
            <Spinner size="tiny" />
          </div>
        ) : filtered.length === 0 ? (
          <div className={styles.empty}>
            <Text>{t('history.empty')}</Text>
          </div>
        ) : (
          filtered.map((c) => {
            const isActive = c.id === activeConversationId;
            const isRenaming = renamingId === c.id;
            return (
              <div
                key={c.id}
                role="button"
                tabIndex={0}
                aria-label={t('history.conversationAria', {
                  title: c.title ?? t('history.untitled'),
                })}
                className={`${styles.row} ${isActive ? styles.rowActive : ''}`}
                onClick={(e) => {
                  if (isRenaming) return;
                  if ((e.target as HTMLElement).closest('[data-row-action]')) return;
                  onSelectConversation(c.id);
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !isRenaming) {
                    onSelectConversation(c.id);
                  }
                }}
              >
                <div className={styles.rowMain}>
                  {isRenaming ? (
                    <Input
                      value={renameDraft}
                      onChange={(_, d) => setRenameDraft(d.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') void commitRename();
                        if (e.key === 'Escape') cancelRename();
                      }}
                      onBlur={() => void commitRename()}
                      autoFocus
                      aria-label={t('history.renameAria')}
                    />
                  ) : (
                    <div className={styles.rowTitle} dir="auto">
                      {c.title ?? t('history.untitled')}
                    </div>
                  )}
                  <div className={styles.rowMeta}>
                    <Badge appearance="outline" size="small">
                      {hostLabel(c.host)}
                    </Badge>
                    <span>{t('history.updated', { time: relativeTime(c.updatedAt) })}</span>
                    <CostBadge cost={c.usageCost} />
                  </div>
                </div>
                <div className={styles.rowActions} data-row-action="">
                  {isRenaming ? (
                    <>
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<Checkmark20Regular />}
                        onClick={() => void commitRename()}
                        aria-label={t('history.saveNameAria')}
                      />
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<DismissCircle20Regular />}
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={cancelRename}
                        aria-label={t('history.cancelRenameAria')}
                      />
                    </>
                  ) : (
                    <>
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<Edit20Regular />}
                        onClick={() => startRename(c)}
                        aria-label={t('history.renameAria')}
                      />
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<Delete20Regular />}
                        onClick={() => void requestDelete(c.id)}
                        aria-label={t('history.deleteAria')}
                      />
                    </>
                  )}
                </div>
              </div>
            );
          })
        )}
      </div>

    </div>
  );
}
