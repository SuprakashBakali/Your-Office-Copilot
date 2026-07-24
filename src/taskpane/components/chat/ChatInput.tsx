import React, { useState, useRef, KeyboardEvent, useEffect } from 'react';
import { makeStyles, tokens, Button, Textarea, Tooltip, Text } from '@fluentui/react-components';
import { SendRegular, StopRegular, KeyboardRegular } from '@fluentui/react-icons';
import { useAppState, useAppDispatch } from '../../store/AppContext';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    width: '100%',
  },
  inputArea: {
    display: 'flex',
    gap: '6px',
    alignItems: 'flex-end',
  },
  textarea: {
    flexGrow: 1,
    minHeight: '40px',
    maxHeight: '120px',
  },
  sendBtn: {
    minWidth: '36px',
    height: '36px',
    borderRadius: '50%',
  },
  hint: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    paddingLeft: '4px',
    paddingRight: '4px',
  },
});

interface ChatInputProps {
  onSend: (text: string) => void;
  disabled?: boolean;
  onCancel?: () => void;
  isStreaming?: boolean;
}

export const ChatInput: React.FC<ChatInputProps> = ({ onSend, disabled, onCancel, isStreaming }) => {
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
          placeholder={isStreaming ? 'Generating response...' : 'Ask Copilot anything...'}
          resize="none"
          disabled={disabled}
          size="medium"
        />
        {isStreaming ? (
          <Tooltip content="Stop generating" relationship="label">
            <Button
              className={classes.sendBtn}
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
      </div>
    </div>
  );
};
