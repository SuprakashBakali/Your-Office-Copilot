import React, { useState } from 'react';
import { makeStyles, tokens } from '@fluentui/react-components';
import { ChevronDownRegular, ChevronRightRegular, BrainCircuitRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    marginBottom: '8px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: '6px',
    backgroundColor: tokens.colorNeutralBackground2,
    overflow: 'hidden',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    padding: '4px 8px',
    width: '100%',
    background: 'none',
    border: 'none',
    cursor: 'pointer',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorBrandForeground1,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground2Hover,
    },
  },
  chevron: {
    display: 'inline-flex',
    alignItems: 'center',
    fontSize: '10px',
  },
  body: {
    padding: '8px 10px',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    maxHeight: '160px',
    overflowY: 'auto',
    fontFamily: 'monospace',
  },
  streaming: {
    display: 'inline-block',
    marginLeft: '4px',
    animationName: {
      '0%, 100%': { opacity: 0.3 },
      '50%': { opacity: 1 },
    },
    animationDuration: '1.4s',
    animationIterationCount: 'infinite',
  },
});

interface ThinkingBlockProps {
  thinking: string;
  isStreaming?: boolean;
}

/**
 * Collapsible reasoning/thinking block — renders the model's chain-of-thought
 * trace (DeepSeek-R1 `<think>`, Claude extended thinking, etc.) in a subdued
 * panel above the assistant's actual answer.
 *
 * Pattern adapted from office-agents' `thinking-block.svelte`.
 */
export const ThinkingBlock: React.FC<ThinkingBlockProps> = ({ thinking, isStreaming }) => {
  const classes = useStyles();
  const [expanded, setExpanded] = useState(false);

  if (!thinking || !thinking.trim()) return null;

  return (
    <div className={classes.container}>
      <button
        type="button"
        className={classes.header}
        onClick={() => setExpanded(e => !e)}
        aria-expanded={expanded}
      >
        <span className={classes.chevron}>
          {expanded ? <ChevronDownRegular /> : <ChevronRightRegular />}
        </span>
        <BrainCircuitRegular />
        <span>thinking</span>
        {isStreaming && <span className={classes.streaming}>...</span>}
      </button>
      {expanded && (
        <div className={classes.body}>
          {thinking}
        </div>
      )}
    </div>
  );
};
