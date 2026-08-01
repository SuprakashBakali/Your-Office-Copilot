import React, { useEffect, useRef, useState } from 'react';
import {
  makeStyles, tokens, Button, Dropdown, Option, Tooltip, Switch, Text,
  Menu, MenuItem, MenuTrigger, MenuPopover, MenuList, Divider,
} from '@fluentui/react-components';
import {
  AddRegular, ArrowExportRegular, ChatRegular, DeleteRegular,
  DocumentTextRegular, CodeRegular, TextTRegular, HistoryRegular,
} from '@fluentui/react-icons';
import { useChat } from '../../hooks/useChat';
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
  // Bottom toolbar — replaces where Tools/Templates used to be
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 8px 6px 8px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
  historyDropdown: {
    flexGrow: 1,
    minWidth: 0,
  },
  toolbarActions: {
    display: 'flex',
    gap: '2px',
    alignItems: 'center',
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

  const handleExport = React.useCallback((format: 'txt' | 'json' | 'markdown') => {
    if (activeConversation) {
      exportConversation(activeConversation.id, format);
    }
  }, [activeConversation, exportConversation]);

  return (
    <div className={classes.container}>
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

      {/* Bottom Toolbar — History, New Chat, Export, Delete */}
      <div className={classes.toolbar}>
        <Tooltip content="Chat history" relationship="label">
          <HistoryRegular style={{ fontSize: '14px', color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
        </Tooltip>
        <Dropdown
          className={classes.historyDropdown}
          value={activeConversation?.title || 'New Chat'}
          onOptionSelect={(_, data) => setActiveConversation(data.optionValue as string)}
          size="small"
        >
          {conversations.map(conv => (
            <Option key={conv.id} value={conv.id}>{conv.title}</Option>
          ))}
        </Dropdown>

        <div className={classes.toolbarActions}>
          <Tooltip content="New Chat" relationship="label">
            <Button
              icon={<AddRegular />}
              appearance="subtle"
              size="small"
              onClick={createConversation}
            />
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
      </div>
    </div>
  );
};
