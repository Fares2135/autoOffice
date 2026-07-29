import React from 'react';
import { makeStyles, tokens, Text } from '@fluentui/react-components';
import { Checkmark12Regular } from '@fluentui/react-icons';
import { useTranslation } from '../i18n/index.ts';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '4px 12px',
    color: tokens.colorNeutralForeground3,
    fontSize: '12px',
  },
});

export function ToolActivity({ toolName }: { toolName: string }) {
  const styles = useStyles();
  const { t } = useTranslation();
  // Inspection is not a lookup, so it gets its own phrasing rather than
  // reading as "looked up: inspect_document".
  const label = toolName === 'inspect_document'
    ? t('code.inspectActivity')
    : t('code.toolActivity', { toolName });
  return (
    <div className={styles.container}>
      <Checkmark12Regular />
      <Text size={200} italic>{label}</Text>
    </div>
  );
}
