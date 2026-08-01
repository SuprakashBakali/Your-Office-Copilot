import React, { useEffect, useRef, useState } from 'react';
import { makeStyles, tokens, Switch, Text } from '@fluentui/react-components';
import { ChatRegular } from '@fluentui/react-icons';
import { UseChatReturn } from '../../hooks/useChat';
import { useAI } from '../../hooks/useAI';
import { useAppState } from '../../store/AppContext';
import { MessageBubble } from './MessageBubble';
import { ChatInput } from './ChatInput';
import { EmptyState } from '../shared/EmptyState';
import { LoadingDots } from '../shared/LoadingDots';
import { useSettings } from '../../hooks/useSettings';

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
    padding: '10px 8px',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  inputArea: {
    padding: '6px 8px 4px 8px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
  contextRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    marginBottom: '4px',
    fontSize: tokens.fontSizeBase200,
  },
  modelInfo: {
    display: 'flex',
    alignItems: 'center',
    marginLeft: 'auto',
    padding: '2px 6px',
    borderRadius: '4px',
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    maxWidth: '130px',
  },
});

interface ChatPanelProps {
  chat: UseChatReturn;
}

export const ChatPanel: React.FC<ChatPanelProps> = ({ chat }) => {
  const classes = useStyles();
  const { host } = useAppState();
  const {
    activeConversation, messages, sendChatMessage,
    createConversation,
  } = chat;
  const { isStreaming, cancelStream } = useAI();
  const messageListRef = useRef<HTMLDivElement>(null);
  const [includeContext, setIncludeContext] = useState(true);
  const { settings } = useSettings();

  // Auto-scroll on new messages
  useEffect(() => {
    if (messageListRef.current) {
      messageListRef.current.scrollTop = messageListRef.current.scrollHeight;
    }
  }, [messages, isStreaming]);

  const handleSend = React.useCallback((text: string) => {
    sendChatMessage(text, includeContext);
  }, [sendChatMessage, includeContext]);

  return (
    <div className={classes.container}>
      {/* Messages — Now takes up maximum vertical screen space! */}
      <div className={classes.messageList} ref={messageListRef}>
        {messages.length === 0 ? (
          <EmptyState
            icon={<ChatRegular />}
            title="Welcome to AI Copilot"
            description={`Chat with your ${host} workbook. Ask questions, analyze data, generate formulas, and more.`}
            actionText="Start a new chat"
            onAction={createConversation}
          />
        ) : (
          messages.filter(m => m.role !== 'system').map(msg => (
            <MessageBubble key={msg.id} message={msg} />
          ))
        )}
        {isStreaming && (
          <div style={{ padding: '8px', display: 'flex', justifyContent: 'flex-start' }}>
            <LoadingDots label="Generating..." />
          </div>
        )}
      </div>

      {/* Input Area */}
      <div className={classes.inputArea}>
        <div className={classes.contextRow}>
          <Switch
            checked={includeContext}
            onChange={(_, data) => setIncludeContext(data.checked)}
            label=""
          />
          <Text size={200} style={{ whiteSpace: 'nowrap' }}>
            {includeContext ? '📎 Context ON' : 'Context OFF'}
          </Text>
          <div className={classes.modelInfo} title={`${settings.activeProvider}/${settings.activeModel}`}>
            {settings.activeModel.split('/').pop()}
          </div>
        </div>
        <ChatInput
          onSend={handleSend}
          disabled={isStreaming}
          onCancel={cancelStream}
          isStreaming={isStreaming}
        />
      </div>
    </div>
  );
};
