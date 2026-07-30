import React from 'react';
import { makeStyles, tokens, Text } from '@fluentui/react-components';
import type { ChatMessage } from '../agent/orchestrator.ts';
import { CodeBlock } from './CodeBlock.tsx';
import { ToolActivity } from './ToolActivity.tsx';
import { ErrorBubble } from './ErrorBubble.tsx';
import { Markdown } from './Markdown.tsx';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    padding: '8px 12px',
    maxWidth: '100%',
  },
  userBubble: {
    alignSelf: 'flex-end',
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    borderRadius: '12px 12px 4px 12px',
    padding: '8px 12px',
    maxWidth: '85%',
    wordBreak: 'break-word',
  },
  assistantBubble: {
    alignSelf: 'flex-start',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: '12px 12px 12px 4px',
    padding: '8px 12px',
    maxWidth: '85%',
    wordBreak: 'break-word',
  },
  messageText: {
    whiteSpace: 'pre-wrap',
    lineHeight: '1.4',
    fontSize: '13px',
  },
});

export function MessageBubble({ message }: { message: ChatMessage }) {
  const styles = useStyles();

  if (message.error) {
    return (
      <div className={styles.container}>
        <ErrorBubble {...message.error} />
      </div>
    );
  }

  if (message.toolActivity) {
    return <ToolActivity toolName={message.toolActivity.toolName} />;
  }

  if (message.codeBlock) {
    return (
      <div className={styles.container}>
        <CodeBlock
          code={message.codeBlock.code}
          status={message.codeBlock.status}
          result={message.codeBlock.result}
        />
      </div>
    );
  }

  if (!message.content) return null;

  // Only assistant replies are markdown. User text stays literal so their own
  // asterisks and underscores are not reinterpreted as formatting.
  if (message.role === 'assistant') {
    return (
      <div className={styles.container}>
        <div className={styles.assistantBubble}>
          <Markdown>{message.content}</Markdown>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.userBubble}>
        <Text className={styles.messageText}>{message.content}</Text>
      </div>
    </div>
  );
}
