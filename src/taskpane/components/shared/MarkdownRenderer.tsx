import React from 'react';
import ReactMarkdown, { type Components } from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { makeStyles, tokens } from '@fluentui/react-components';
import { CodeBlock } from './CodeBlock';

const useStyles = makeStyles({
  container: {
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    color: tokens.colorNeutralForeground1,
    '& p': {
      marginBottom: '12px',
    },
    '& a': {
      color: tokens.colorBrandForeground1,
      textDecoration: 'none',
    },
    '& a:hover': {
      textDecoration: 'underline',
    },
    '& table': {
      width: '100%',
      borderCollapse: 'collapse',
      marginBottom: '12px',
    },
    '& th, & td': {
      border: `1px solid ${tokens.colorNeutralStroke1}`,
      padding: '6px',
    },
    '& th': {
      backgroundColor: tokens.colorNeutralBackground3,
      fontWeight: tokens.fontWeightSemibold,
    },
  },
});

interface MarkdownRendererProps {
  content: string;
}

// react-markdown v9 removed the `inline` prop on the `code` component.
// We detect inline code by checking whether `className` contains
// `language-*` (block code always has a language class) and whether the
// node's position spans a single line.
const components: Components = {
  code({ node, className, children, ...props }) {
    const match = /language-(\w+)/.exec(className || '');
    const isBlock = !!match || (node?.position && node.position.start.line !== node.position.end.line);
    if (isBlock && match) {
      return (
        <CodeBlock
          language={match[1]}
          value={String(children).replace(/\n$/, '')}
        />
      );
    }
    return (
      <code
        className={className}
        style={{
          backgroundColor: tokens.colorNeutralBackground3,
          padding: '2px 4px',
          borderRadius: '4px',
          fontFamily: 'monospace',
        }}
        {...props}
      >
        {children}
      </code>
    );
  },
};

export const MarkdownRenderer: React.FC<MarkdownRendererProps> = React.memo(({ content }) => {
  const classes = useStyles();

  return (
    <div className={classes.container}>
      <ReactMarkdown remarkPlugins={[remarkGfm]} components={components}>
        {content}
      </ReactMarkdown>
    </div>
  );
});
