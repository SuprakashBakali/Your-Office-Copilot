import React from 'react';
import { Highlight, type PrismTheme } from 'prism-react-renderer';
import { makeStyles, tokens } from '@fluentui/react-components';
import { CopyButton } from './CopyButton';

const useStyles = makeStyles({
  container: {
    margin: '12px 0',
    borderRadius: '8px',
    overflow: 'hidden',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '4px 12px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
});

// OneDark-style theme for prism-react-renderer. Keeps the dark code panel
// regardless of app light/dark mode (consistent with most IDEs).
const oneDarkTheme: PrismTheme = {
  plain: { color: '#abb2bf', backgroundColor: '#1E1E1E' },
  styles: [
    { types: ['comment', 'prolog', 'doctype', 'cdata'], style: { color: '#5c6370', fontStyle: 'italic' } },
    { types: ['punctuation'], style: { color: '#abb2bf' } },
    { types: ['property', 'tag', 'constant', 'symbol', 'deleted'], style: { color: '#e06c75' } },
    { types: ['boolean', 'number'], style: { color: '#d19a66' } },
    { types: ['selector', 'attr-name', 'string', 'char', 'builtin', 'inserted'], style: { color: '#98c379' } },
    { types: ['operator', 'entity', 'url'], style: { color: '#56b6c2' } },
    { types: ['atrule', 'attr-value', 'keyword'], style: { color: '#c678dd' } },
    { types: ['function', 'class-name'], style: { color: '#61afef' } },
    { types: ['regex', 'important', 'variable'], style: { color: '#e06c75' } },
  ],
};

interface CodeBlockProps {
  language: string;
  value: string;
}

export const CodeBlock: React.FC<CodeBlockProps> = React.memo(({ language, value }) => {
  const classes = useStyles();

  return (
    <div className={classes.container}>
      <div className={classes.header}>
        <span>{language}</span>
        <CopyButton text={value} />
      </div>
      <Highlight theme={oneDarkTheme} code={value} language={language as any}>
        {({ className, style, tokens, getLineProps, getTokenProps }) => (
          <pre
            className={className}
            style={{
              ...style,
              margin: 0,
              padding: '12px',
              fontSize: '13px',
              overflowX: 'auto',
            }}
          >
            {tokens.map((line, i) => {
              const lineProps = getLineProps({ line });
              return (
                <div key={i} {...lineProps}>
                  {line.map((token, key) => (
                    <span key={key} {...getTokenProps({ token })} />
                  ))}
                </div>
              );
            })}
          </pre>
        )}
      </Highlight>
    </div>
  );
});
