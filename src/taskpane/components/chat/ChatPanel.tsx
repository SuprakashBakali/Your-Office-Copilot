import React, { useEffect, useRef } from 'react';
import { makeStyles, tokens } from '@fluentui/react-components';
import { ChatRegular } from '@fluentui/react-icons';
import { UseChatReturn } from '../../hooks/useChat';
import { useAppState } from '../../store/AppContext';
import { MessageBubble } from './MessageBubble';
import { ChatInput } from './ChatInput';
import { EmptyState } from '../shared/EmptyState';
import { LoadingDots } from '../shared/LoadingDots';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    width: '100%',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  messageList: {
    flexGrow: 1,
    overflowY: 'auto',
    padding: '12px 12px',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  inputArea: {
    padding: '8px 12px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
});

interface ChatPanelProps {
  chat: UseChatReturn;
  webSearchEnabled?: boolean;
}

export const ChatPanel: React.FC<ChatPanelProps> = ({ chat, webSearchEnabled = false }) => {
  const classes = useStyles();
  const { host } = useAppState();
  const {
    messages, sendChatMessage,
    createConversation, includeContext,
    isStreaming, cancelStream,
  } = chat;
  // useAI's error state is not exposed through useChat — but ChatInput can
  // show a generic error indicator if the last message starts with ⚠️.
  const lastMessage = messages[messages.length - 1];
  const hasError = lastMessage?.role === 'assistant' && (lastMessage.content || '').startsWith('⚠️');
  const messageListRef = useRef<HTMLDivElement>(null);
  const isAtBottomRef = useRef(true);

  // Auto-scroll on new messages — but only if the user is already near the
  // bottom (within 80px). This prevents the "scroll trap" where the user
  // tries to scroll up to read earlier messages but gets forced back down
  // on every streamed chunk.
  useEffect(() => {
    const el = messageListRef.current;
    if (el && isAtBottomRef.current) {
      el.scrollTop = el.scrollHeight;
    }
  }, [messages, isStreaming]);

  const handleScroll = () => {
    const el = messageListRef.current;
    if (!el) return;
    isAtBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 80;
  };

  const handleSend = React.useCallback((text: string, attachments?: any[]) => {
    sendChatMessage(text, includeContext, webSearchEnabled, attachments);
  }, [sendChatMessage, includeContext, webSearchEnabled]);

  return (
    <div className={classes.container}>
      {/* Messages — Maximum vertical screen space */}
      <div className={classes.messageList} ref={messageListRef} onScroll={handleScroll}>
        {messages.length === 0 ? (
          <EmptyState
            icon={<ChatRegular />}
            title="Welcome to AI Copilot"
            description={`Chat with your ${host} workbook. Ask questions, analyze data, generate formulas, and more.`}
            actionText="Start a new chat"
            onAction={createConversation}
          />
        ) : (
          messages.filter(m => m.role !== 'system' && m.role !== 'compaction_summary').map(msg => (
            <MessageBubble key={msg.id} message={msg} />
          ))
        )}
        {isStreaming && (
          <div style={{ padding: '8px', display: 'flex', justifyContent: 'flex-start' }}>
            <LoadingDots label="Generating..." />
          </div>
        )}
      </div>

      {/* Input Area — Clean & Unified Padding (No Model Card / Context Switch) */}
      <div className={classes.inputArea}>
        <ChatInput
          onSend={handleSend}
          disabled={isStreaming}
          onCancel={cancelStream}
          isStreaming={isStreaming}
          error={hasError && !isStreaming ? 'last message had an error' : undefined}
        />
      </div>
    </div>
  );
};
