import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  makeStyles, tokens, Button, Dropdown, Option, Tooltip, Switch, Text,
  Badge, Menu, MenuItem, MenuTrigger, MenuPopover, MenuList,
} from '@fluentui/react-components';
import {
  AddRegular, ArrowExportRegular, ChatRegular, DeleteRegular,
  DocumentTextRegular, CodeRegular, TextTRegular,
} from '@fluentui/react-icons';
import { useChat } from '../../hooks/useChat';
import { useAI } from '../../hooks/useAI';
import { useAppState } from '../../store/AppContext';
import { MessageBubble } from './MessageBubble';
import { ChatInput } from './ChatInput';
import { EmptyState } from '../shared/EmptyState';
import { LoadingDots } from '../shared/LoadingDots';
import { loadSettings } from '../../utils/storage';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    width: '100%',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    padding: '6px 10px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    gap: '6px',
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
  conversationSelect: {
    flexGrow: 1,
    minWidth: 0,
  },
  messageList: {
    flexGrow: 1,
    overflowY: 'auto',
    padding: '12px',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  inputArea: {
    padding: '10px 12px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
  contextToggle: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    marginBottom: '8px',
    fontSize: tokens.fontSizeBase200,
  },
  modelInfo: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    padding: '2px 6px',
    borderRadius: '4px',
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    flexShrink: 0,
  },
});

export const ChatPanel: React.FC = () => {
  const classes = useStyles();
  const { host } = useAppState();
  const {
    conversations, activeConversation, messages, sendChatMessage,
    createConversation, setActiveConversation, deleteConversation, exportConversation,
  } = useChat(host);
  const { isStreaming, cancelStream } = useAI();
  const messageListRef = useRef<HTMLDivElement>(null);
  const [includeContext, setIncludeContext] = useState(true);
  const settings = useMemo(() => loadSettings(), [messages.length]);

  // Auto-scroll on new messages
  useEffect(() => {
    if (messageListRef.current) {
      messageListRef.current.scrollTop = messageListRef.current.scrollHeight;
    }
  }, [messages, isStreaming]);

  const handleSend = (text: string) => {
    sendChatMessage(text, includeContext);
  };

  const handleExport = (format: 'txt' | 'json' | 'markdown') => {
    if (activeConversation) {
      exportConversation(activeConversation.id, format);
    }
  };

  return (
    <div className={classes.container}>
      {/* Header with conversation management */}
      <div className={classes.header}>
        <Dropdown
          className={classes.conversationSelect}
          value={activeConversation?.title || 'New Chat'}
          onOptionSelect={(_, data) => setActiveConversation(data.optionValue as string)}
          size="small"
        >
          {conversations.map(conv => (
            <Option key={conv.id} value={conv.id}>{conv.title}</Option>
          ))}
        </Dropdown>

        <Tooltip content="New Chat" relationship="label">
          <Button icon={<AddRegular />} appearance="subtle" size="small" onClick={createConversation} />
        </Tooltip>

        <Menu>
          <MenuTrigger>
            <Tooltip content="Export" relationship="label">
              <Button icon={<ArrowExportRegular />} appearance="subtle" size="small" />
            </Tooltip>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem icon={<DocumentTextRegular />} onClick={() => handleExport('markdown')}>
                Export as Markdown
              </MenuItem>
              <MenuItem icon={<CodeRegular />} onClick={() => handleExport('json')}>
                Export as JSON
              </MenuItem>
              <MenuItem icon={<TextTRegular />} onClick={() => handleExport('txt')}>
                Export as Text
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>

        {activeConversation && (
          <Tooltip content="Delete conversation" relationship="label">
            <Button
              icon={<DeleteRegular />}
              appearance="subtle"
              size="small"
              onClick={() => deleteConversation(activeConversation.id)}
            />
          </Tooltip>
        )}
      </div>

      {/* Messages */}
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
        {isStreaming && messages.length > 0 && messages[messages.length - 1]?.role !== 'assistant' && (
          <div style={{ padding: '8px', display: 'flex', justifyContent: 'flex-start' }}>
            <LoadingDots label="Generating..." />
          </div>
        )}
      </div>

      {/* Input Area */}
      <div className={classes.inputArea}>
        <div className={classes.contextToggle}>
          <Switch
            checked={includeContext}
            onChange={(_, data) => setIncludeContext(data.checked)}
            label=""
          />
          <Text size={200}>
            {includeContext ? '📎 Workbook context ON' : 'Workbook context OFF'}
          </Text>
          <div className={classes.modelInfo}>
            {settings.activeProvider}/{settings.activeModel.split('/').pop()}
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
