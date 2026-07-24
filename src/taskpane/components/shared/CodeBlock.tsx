import React from 'react';
import { Prism as SyntaxHighlighter } from 'react-syntax-highlighter';
import { oneDark } from 'react-syntax-highlighter/dist/esm/styles/prism';
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

interface CodeBlockProps {
  language: string;
  value: string;
}

export const CodeBlock: React.FC<CodeBlockProps> = ({ language, value }) => {
  const classes = useStyles();

  return (
    <div className={classes.container}>
      <div className={classes.header}>
        <span>{language}</span>
        <CopyButton text={value} />
      </div>
      <SyntaxHighlighter
        language={language}
        style={oneDark as any}
        customStyle={{
          margin: 0,
          padding: '12px',
          fontSize: '13px',
          backgroundColor: '#1E1E1E' // Force dark bg
        }}
      >
        {value}
      </SyntaxHighlighter>
    </div>
  );
};
