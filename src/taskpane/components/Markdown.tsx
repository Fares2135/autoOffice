import React from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { makeStyles, tokens } from '@fluentui/react-components';

// Assistant replies are markdown — headings, lists, tables, inline code — and
// used to render as literal asterisks and pipes. Raw HTML stays disabled
// (react-markdown's default), so model output cannot inject markup.
const useStyles = makeStyles({
  root: {
    fontSize: '13px',
    lineHeight: '1.5',
    // dir="auto" per block: an Arabic paragraph and an English one in the same
    // reply each get the right direction.
    '& > *:first-child': { marginTop: 0 },
    '& > *:last-child': { marginBottom: 0 },
    '& p': { margin: '0 0 8px' },
    '& ul, & ol': { margin: '0 0 8px', paddingInlineStart: '20px' },
    '& li': { margin: '2px 0' },
    '& h1, & h2, & h3, & h4': {
      margin: '12px 0 6px',
      fontWeight: 600,
      lineHeight: '1.3',
    },
    '& h1': { fontSize: '16px' },
    '& h2': { fontSize: '15px' },
    '& h3': { fontSize: '14px' },
    '& h4': { fontSize: '13px' },
    '& code': {
      fontFamily: tokens.fontFamilyMonospace,
      fontSize: '12px',
      backgroundColor: tokens.colorNeutralBackground4,
      borderRadius: tokens.borderRadiusSmall,
      padding: '1px 4px',
    },
    '& pre': {
      margin: '0 0 8px',
      padding: '8px 10px',
      overflowX: 'auto',
      backgroundColor: tokens.colorNeutralBackground4,
      borderRadius: tokens.borderRadiusMedium,
      // Code is left-to-right even inside an RTL reply.
      direction: 'ltr',
      textAlign: 'left',
    },
    '& pre code': {
      padding: 0,
      backgroundColor: 'transparent',
      whiteSpace: 'pre',
    },
    '& blockquote': {
      margin: '0 0 8px',
      paddingInlineStart: '10px',
      borderInlineStartWidth: '3px',
      borderInlineStartStyle: 'solid',
      borderInlineStartColor: tokens.colorNeutralStroke1,
      color: tokens.colorNeutralForeground2,
    },
    '& a': { color: tokens.colorBrandForegroundLink },
    '& hr': {
      height: '1px',
      backgroundColor: tokens.colorNeutralStroke2,
      margin: '10px 0',
    },
    '& table': { borderCollapse: 'collapse', fontSize: '12px' },
    // Four single-side shorthands: griffel's types reject the all-sides ones.
    '& th, & td': {
      borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
      borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
      borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
      borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
      padding: '3px 6px',
      textAlign: 'start',
    },
    '& th': { backgroundColor: tokens.colorNeutralBackground4, fontWeight: 600 },
  },
  // Wide tables scroll inside the bubble instead of stretching the task pane.
  tableScroll: {
    overflowX: 'auto',
    margin: '0 0 8px',
  },
});

export function Markdown({ children }: { children: string }) {
  const styles = useStyles();
  return (
    <div className={styles.root} dir="auto">
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        components={{
          a: ({ children: c, ...props }) => (
            <a {...props} target="_blank" rel="noopener noreferrer">{c}</a>
          ),
          table: ({ children: c }) => (
            <div className={styles.tableScroll}>
              <table>{c}</table>
            </div>
          ),
        }}
      >
        {children}
      </ReactMarkdown>
    </div>
  );
}
