import React from 'react';
import { makeStyles, tokens } from '@fluentui/react-components';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { CodeBlock } from '../CodeBlock.tsx';

const useStyles = makeStyles({
  markdown: {
    minWidth: 0,
    lineHeight: 1.58,
    overflowWrap: 'anywhere',
    '& > :first-child': { marginTop: 0 },
    '& > :last-child': { marginBottom: 0 },
    '& p': { marginTop: '0.45em', marginBottom: '0.45em' },
    '& h1, & h2, & h3': {
      marginTop: '0.9em',
      marginBottom: '0.35em',
      lineHeight: 1.25,
    },
    '& ul, & ol': {
      marginTop: '0.45em',
      marginBottom: '0.45em',
      paddingInlineStart: '1.4em',
    },
    '& li': { marginBlock: '0.2em' },
    '& blockquote': {
      marginInline: 0,
      paddingInlineStart: '12px',
      borderLeft: `3px solid ${tokens.colorBrandStroke1}`,
      color: tokens.colorNeutralForeground2,
    },
    '& a': {
      color: tokens.colorBrandForegroundLink,
      textDecorationThickness: '1px',
      textUnderlineOffset: '2px',
    },
    '& code': {
      direction: 'ltr',
      unicodeBidi: 'isolate',
      fontFamily: 'Consolas, "SFMono-Regular", monospace',
      fontSize: '0.88em',
      padding: '1px 4px',
      borderRadius: '4px',
      backgroundColor: tokens.colorNeutralBackground5,
    },
    '& table': {
      width: '100%',
      borderCollapse: 'collapse',
      marginBlock: '8px',
      fontSize: '0.92em',
    },
    '& th, & td': {
      padding: '6px 8px',
      border: `1px solid ${tokens.colorNeutralStroke2}`,
      textAlign: 'start',
    },
    '& th': { backgroundColor: tokens.colorNeutralBackground4 },
    '& pre': { margin: 0 },
  },
  codeBlock: {
    marginBlock: '8px',
    overflowX: 'auto',
    borderRadius: '8px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground5,
    '& pre': {
      margin: 0,
      padding: '10px',
      overflowX: 'auto',
    },
  },
});

export function TextPart({ part }: { part: { text: string } }) {
  const styles = useStyles();
  return (
    <div className={styles.markdown} dir="auto">
      <ReactMarkdown
        skipHtml
        remarkPlugins={[remarkGfm]}
        components={{
          p: ({ children }) => <p dir="auto">{children}</p>,
          li: ({ children }) => <li dir="auto">{children}</li>,
          h1: ({ children }) => <h2 dir="auto">{children}</h2>,
          h2: ({ children }) => <h3 dir="auto">{children}</h3>,
          h3: ({ children }) => <h4 dir="auto">{children}</h4>,
          a: ({ href, children }) => (
            <a href={href} target="_blank" rel="noopener noreferrer">
              {children}
            </a>
          ),
          pre: ({ children }) => <>{children}</>,
          code: ({ className, children }) => {
            const text = String(children).replace(/\n$/, '');
            const language = /language-(typescript|ts)/.test(className ?? '')
              ? 'typescript'
              : 'javascript';
            const block = Boolean(className) || text.includes('\n');
            return block ? (
              <div className={styles.codeBlock} dir="ltr">
                <CodeBlock code={text} lang={language} />
              </div>
            ) : (
              <code dir="ltr">{children}</code>
            );
          },
        }}
      >
        {part.text}
      </ReactMarkdown>
    </div>
  );
}
