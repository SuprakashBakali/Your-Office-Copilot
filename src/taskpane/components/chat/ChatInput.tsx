import React, { useState, useRef, KeyboardEvent, useEffect } from 'react';
import { makeStyles, tokens, Button, Textarea, Tooltip, Text, Spinner } from '@fluentui/react-components';
import { SendRegular, StopRegular, KeyboardRegular } from '@fluentui/react-icons';
import { useAppState, useAppDispatch } from '../../store/AppContext';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    width: '100%',
  },
  inputArea: {
    display: 'flex',
    gap: '6px',
    alignItems: 'flex-end',
  },
  textarea: {
    flexGrow: 1,
    minHeight: '36px',
    maxHeight: '120px',
  },
  sendBtn: {
    minWidth: '36px',
    height: '36px',
    borderRadius: '50%',
    transition: 'all 0.2s ease',
  },
  // When streaming, the send button gets a subtle pulsing border so the
  // user can see at a glance that the LLM is still generating.
  sendBtnStreaming: {
    animationName: {
      '0%, 100%': { boxShadow: '0 0 0 0 rgba(118, 185, 0, 0.4)' },
      '50%': { boxShadow: '0 0 0 6px rgba(118, 185, 0, 0)' },
    },
    animationDuration: '1.5s',
    animationIterationCount: 'infinite',
  },
  hint: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    paddingLeft: '4px',
    paddingRight: '4px',
    marginTop: '-2px',
  },
  status: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    fontSize: tokens.fontSizeBase100,
  },
  statusDot: {
    width: '6px',
    height: '6px',
    borderRadius: '50%',
    display: 'inline-block',
  },
  statusIdle: {
    backgroundColor: tokens.colorNeutralForeground4,
  },
  statusStreaming: {
    backgroundColor: '#76B900',
    animationName: {
      '0%, 100%': { opacity: 0.4 },
      '50%': { opacity: 1 },
    },
    animationDuration: '1.4s',
    animationIterationCount: 'infinite',
  },
  statusError: {
    backgroundColor: tokens.colorPaletteRedForeground1,
  },
  statusText: {
    color: tokens.colorNeutralForeground4,
  },
  statusTextActive: {
    color: '#76B900',
    fontWeight: 600,
  },
});

interface ChatInputProps {
  onSend: (text: string) => void;
  disabled?: boolean;
  onCancel?: () => void;
  isStreaming?: boolean;
  error?: string | null;
}

export const ChatInput: React.FC<ChatInputProps> = ({ onSend, disabled, onCancel, isStreaming, error }) => {
  const classes = useStyles();
  const [text, setText] = useState('');
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const { pendingPrompt } = useAppState();
  const dispatch = useAppDispatch();

  useEffect(() => {
    if (pendingPrompt) {
      setText(pendingPrompt);
      dispatch({ type: 'SET_PENDING_PROMPT', payload: null });
      if (textareaRef.current) {
        textareaRef.current.focus();
      }
    }
  }, [pendingPrompt, dispatch]);

  const handleSend = () => {
    if (text.trim() && !disabled) {
      onSend(text.trim());
      setText('');
    }
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  return (
    <div className={classes.container}>
      <div className={classes.inputArea}>
        <Textarea
          ref={textareaRef}
          className={classes.textarea}
          value={text}
          onChange={(_, data) => setText(data.value)}
          onKeyDown={handleKeyDown}
          placeholder={isStreaming ? 'Copilot is responding...' : 'Ask Copilot anything...'}
          resize="none"
          disabled={disabled}
          size="medium"
        />
        {isStreaming ? (
          <Tooltip content="Stop generating" relationship="label">
            <Button
              className={`${classes.sendBtn} ${classes.sendBtnStreaming}`}
              icon={<StopRegular />}
              appearance="secondary"
              shape="circular"
              onClick={onCancel}
            />
          </Tooltip>
        ) : (
          <Tooltip content="Send (Enter)" relationship="label">
            <Button
              className={classes.sendBtn}
              icon={<SendRegular />}
              appearance="primary"
              shape="circular"
              onClick={handleSend}
              disabled={!text.trim()}
            />
          </Tooltip>
        )}
      </div>
      <div className={classes.hint}>
        <Text size={100} style={{ color: tokens.colorNeutralForeground4 }}>
          <KeyboardRegular style={{ fontSize: '10px', marginRight: '2px' }} />
          Enter to send · Shift+Enter for new line
        </Text>
        {/* Status indicator — shows idle / generating / error state */}
        <div className={classes.status}>
          {error ? (
            <>
              <span className={`${classes.statusDot} ${classes.statusError}`} />
              <span className={`${classes.statusText}`} style={{ color: tokens.colorPaletteRedForeground1 }}>Error</span>
            </>
          ) : isStreaming ? (
            <>
              <Spinner size="extra-tiny" />
              <span className={`${classes.statusText} ${classes.statusTextActive}`}>Generating...</span>
            </>
          ) : (
            <>
              <span className={`${classes.statusDot} ${classes.statusIdle}`} />
              <span className={classes.statusText}>Ready</span>
            </>
          )}
        </div>
      </div>
    </div>
  );
};
