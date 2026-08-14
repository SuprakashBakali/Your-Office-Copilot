import React, { useState, useRef, KeyboardEvent, useEffect } from 'react';
import { makeStyles, tokens, Button, Textarea, Tooltip, Text, Spinner } from '@fluentui/react-components';
import { SendRegular, StopRegular, KeyboardRegular, AttachRegular, DismissRegular } from '@fluentui/react-icons';
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
  attachBtn: {
    minWidth: '36px',
    height: '36px',
    borderRadius: '50%',
    flexShrink: 0,
  },
  sendBtn: {
    minWidth: '36px',
    height: '36px',
    borderRadius: '50%',
    transition: 'all 0.2s ease',
    flexShrink: 0,
  },
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
  statusDone: {
    backgroundColor: '#76B900',
  },
  statusText: {
    color: tokens.colorNeutralForeground4,
  },
  statusTextActive: {
    color: '#76B900',
    fontWeight: 600,
  },
  statusTextDone: {
    color: '#76B900',
    fontWeight: 600,
  },
  statusTextError: {
    color: tokens.colorPaletteRedForeground1,
    fontWeight: 600,
  },
  attachments: {
    display: 'flex',
    gap: '4px',
    flexWrap: 'wrap',
    marginBottom: '4px',
  },
  attachmentChip: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    padding: '2px 8px',
    borderRadius: '12px',
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground2,
  },
  hiddenInput: {
    display: 'none',
  },
});

export interface Attachment {
  name: string;
  type: string;       // MIME type
  dataUrl: string;    // base64 data URL
  isImage: boolean;
}

interface ChatInputProps {
  onSend: (text: string, attachments?: Attachment[]) => void;
  disabled?: boolean;
  onCancel?: () => void;
  isStreaming?: boolean;
  error?: string | null;
}

export const ChatInput: React.FC<ChatInputProps> = ({ onSend, disabled, onCancel, isStreaming, error }) => {
  const classes = useStyles();
  const [text, setText] = useState('');
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [justFinished, setJustFinished] = useState(false);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const prevStreamingRef = useRef(false);
  const { pendingPrompt } = useAppState();
  const dispatch = useAppDispatch();

  // Detect transition from streaming → idle to show "Done" confirmation
  useEffect(() => {
    if (prevStreamingRef.current && !isStreaming) {
      setJustFinished(true);
      const timer = setTimeout(() => setJustFinished(false), 3000);
      return () => clearTimeout(timer);
    }
    prevStreamingRef.current = !!isStreaming;
  }, [isStreaming]);

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
      onSend(text.trim(), attachments.length > 0 ? attachments : undefined);
      setText('');
      setAttachments([]);
    }
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;
    const newAttachments: Attachment[] = [];
    for (const file of Array.from(files)) {
      if (file.size > 10 * 1024 * 1024) {
        dispatch({ type: 'SET_NOTIFICATION', payload: { type: 'warning', message: `File "${file.name}" is too large (max 10MB)` } });
        continue;
      }
      const dataUrl = await fileToDataUrl(file);
      newAttachments.push({
        name: file.name,
        type: file.type,
        dataUrl,
        isImage: file.type.startsWith('image/'),
      });
    }
    setAttachments(prev => [...prev, ...newAttachments]);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const removeAttachment = (idx: number) => {
    setAttachments(prev => prev.filter((_, i) => i !== idx));
  };

  return (
    <div className={classes.container}>
      {attachments.length > 0 && (
        <div className={classes.attachments}>
          {attachments.map((att, idx) => (
            <div key={idx} className={classes.attachmentChip}>
              <AttachRegular style={{ fontSize: '10px' }} />
              <span>{att.name}</span>
              <DismissRegular
                style={{ fontSize: '10px', cursor: 'pointer', marginLeft: '2px' }}
                onClick={() => removeAttachment(idx)}
              />
            </div>
          ))}
        </div>
      )}
      <input
        ref={fileInputRef}
        type="file"
        multiple
        accept="image/*,application/pdf,.xlsx,.xls,.docx,.doc,.pptx,.ppt,.csv,.txt"
        onChange={handleFileSelect}
        className={classes.hiddenInput}
      />
      <div className={classes.inputArea}>
        <Tooltip content="Attach image/PDF/file" relationship="label">
          <Button
            className={classes.attachBtn}
            icon={<AttachRegular />}
            appearance="subtle"
            shape="circular"
            onClick={() => fileInputRef.current?.click()}
            disabled={disabled}
            size="medium"
          />
        </Tooltip>
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
              disabled={!text.trim() && attachments.length === 0}
            />
          </Tooltip>
        )}
      </div>
      <div className={classes.hint}>
        <Text size={100} style={{ color: tokens.colorNeutralForeground4 }}>
          <KeyboardRegular style={{ fontSize: '10px', marginRight: '2px' }} />
          Enter to send · Shift+Enter for new line
        </Text>
        <div className={classes.status}>
          {error ? (
            <>
              <span className={`${classes.statusDot} ${classes.statusError}`} />
              <span className={`${classes.statusText} ${classes.statusTextError}`}>Error</span>
            </>
          ) : isStreaming ? (
            <>
              <Spinner size="extra-tiny" />
              <span className={`${classes.statusText} ${classes.statusTextActive}`}>Generating...</span>
            </>
          ) : justFinished ? (
            <>
              <span className={`${classes.statusDot} ${classes.statusDone}`} />
              <span className={`${classes.statusText} ${classes.statusTextDone}`}>✓ Done</span>
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

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}
