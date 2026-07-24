import React from 'react';
import { makeStyles, tokens, Badge, Button, Tooltip } from '@fluentui/react-components';
import { PersonRegular, BotRegular, ArrowSyncRegular, TableSimpleRegular } from '@fluentui/react-icons';
import { MarkdownRenderer } from '../shared/MarkdownRenderer';
import { CopyButton } from '../shared/CopyButton';
import { formatDateTime } from '../../utils/formatters';
import { ChatMessage } from '../../types';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    marginBottom: '4px',
    position: 'relative',
    animationName: {
      '0%': { opacity: 0, transform: 'translateY(8px)' },
      '100%': { opacity: 1, transform: 'translateY(0)' },
    },
    animationDuration: '0.3s',
    animationTimingFunction: 'ease-out',
    animationFillMode: 'forwards',
  },
  userContainer: {
    alignItems: 'flex-end',
  },
  assistantContainer: {
    alignItems: 'flex-start',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    marginBottom: '4px',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  bubble: {
    padding: '12px 14px',
    borderRadius: '12px',
    maxWidth: '95%',
    position: 'relative',
    transition: 'background-color 0.2s ease',
  },
  userBubble: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    borderBottomRightRadius: '4px',
  },
  assistantBubble: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
    borderBottomLeftRadius: '4px',
  },
  errorBubble: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
    borderBottomLeftRadius: '4px',
  },
  actions: {
    display: 'flex',
    gap: '2px',
    marginTop: '6px',
    opacity: 0,
    transition: 'opacity 0.2s ease',
    ':hover': {
      opacity: 1,
    },
  },
  actionsVisible: {
    opacity: 1,
  },
  contextBadge: {
    marginTop: '6px',
  },
  bubbleWrapper: {
    ':hover .msg-actions': {
      opacity: '1 !important' as any,
    },
  },
  modelBadge: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground4,
    marginLeft: '4px',
  },
});

interface MessageBubbleProps {
  message: ChatMessage;
  onRegenerate?: () => void;
  onApplyToExcel?: (content: string) => void;
}

export const MessageBubble: React.FC<MessageBubbleProps> = React.memo(({ message, onRegenerate, onApplyToExcel }) => {
  const classes = useStyles();
  const isUser = message.role === 'user';
  const isError = message.content.startsWith('⚠️');
  const isEmpty = !message.content && !isUser;

  return (
    <div className={`${classes.container} ${isUser ? classes.userContainer : classes.assistantContainer}`}>
      <div className={classes.header}>
        {isUser ? <PersonRegular /> : <BotRegular />}
        <span>{isUser ? 'You' : 'Copilot'}</span>
        {message.model && !isUser && (
          <span className={classes.modelBadge}>({message.model.split('/').pop()})</span>
        )}
      </div>
      <div className={classes.bubbleWrapper}>
        <div className={`${classes.bubble} ${isUser ? classes.userBubble : isError ? classes.errorBubble : classes.assistantBubble}`}>
          {isEmpty ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span className="loading-pulse" style={{
                display: 'inline-block',
                width: '8px', height: '8px',
                borderRadius: '50%',
                backgroundColor: tokens.colorBrandBackground,
                animation: 'pulse 1.4s infinite ease-in-out',
              }} />
              <span style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>Thinking...</span>
            </div>
          ) : isUser ? (
            <div style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{message.content}</div>
          ) : (
            <MarkdownRenderer content={message.content} />
          )}
        </div>
        {!isUser && message.content && (
          <div className={`${classes.actions} msg-actions`}>
            <CopyButton text={message.content} />
            {onRegenerate && (
              <Tooltip content="Regenerate response" relationship="label">
                <Button
                  appearance="transparent"
                  size="small"
                  icon={<ArrowSyncRegular />}
                  onClick={onRegenerate}
                />
              </Tooltip>
            )}
            {onApplyToExcel && message.content.includes('```') && (
              <Tooltip content="Apply to Excel" relationship="label">
                <Button
                  appearance="transparent"
                  size="small"
                  icon={<TableSimpleRegular />}
                  onClick={() => onApplyToExcel(message.content)}
                />
              </Tooltip>
            )}
          </div>
        )}
        {message.contextIncluded && (
          <Badge className={classes.contextBadge} appearance="outline" color="informative" size="small">
            📎 Workbook context included
          </Badge>
        )}
      </div>
    </div>
  );
});
