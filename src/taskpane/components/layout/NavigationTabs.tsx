import React, { useState, useMemo } from 'react';
import {
  makeStyles, tokens, Button, Tooltip,
  Menu, MenuItem, MenuTrigger, MenuPopover, MenuList,
  Input,
} from '@fluentui/react-components';
import {
  Chat20Regular, Settings20Regular, HistoryRegular, AddRegular,
  DeleteRegular, DocumentTextRegular, SearchRegular, GlobeRegular,
} from '@fluentui/react-icons';
import { useAppState, useAppDispatch } from '../../store/AppContext';
import { UseChatReturn } from '../../hooks/useChat';
import { PromptTemplatesPanel } from '../shared/PromptTemplatesPanel';

const useStyles = makeStyles({
  nav: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '8px 12px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    flexShrink: 0,
    minHeight: '40px',
  },
  leftSection: {
    display: 'flex',
    alignItems: 'center',
    flexShrink: 0,
  },
  middleSection: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    justifyContent: 'center',
    flexGrow: 1,
  },
  rightSection: {
    display: 'flex',
    alignItems: 'center',
    flexShrink: 0,
  },
  historySearch: {
    margin: '4px 8px',
    width: '220px',
  },
  historyItem: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    minWidth: '220px',
  },
  historyItemTitle: {
    fontSize: tokens.fontSizeBase200,
  },
  historyItemMeta: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
});

interface NavigationTabsProps {
  chat: UseChatReturn;
  /** Whether web search is enabled — controlled here so the toggle stays in sync. */
  webSearchEnabled?: boolean;
  onToggleWebSearch?: () => void;
}

export const NavigationTabs: React.FC<NavigationTabsProps> = ({
  chat,
  webSearchEnabled = false,
  onToggleWebSearch,
}) => {
  const classes = useStyles();
  const state = useAppState();
  const dispatch = useAppDispatch();
  const [historyQuery, setHistoryQuery] = useState('');

  const filteredConversations = useMemo(() => {
    if (!historyQuery.trim()) return chat.conversations;
    const q = historyQuery.toLowerCase();
    return chat.conversations.filter(c =>
      c.title.toLowerCase().includes(q) ||
      c.messages.some(m => m.content.toLowerCase().includes(q)),
    );
  }, [chat.conversations, historyQuery]);

  const handlePickTemplate = (prompt: string) => {
    dispatch({ type: 'SET_PENDING_PROMPT', payload: prompt });
  };

  return (
    <div className={classes.nav}>
      {/* Left: Chat Icon Button */}
      <div className={classes.leftSection}>
        <Tooltip content="Chat" relationship="label">
          <Button
            appearance={state.view === 'chat' ? 'subtle' : 'transparent'}
            icon={<Chat20Regular />}
            onClick={() => dispatch({ type: 'SET_VIEW', payload: 'chat' })}
            size="small"
          />
        </Tooltip>
      </div>

      {/* Middle: Context Toggle, Templates, Web Search, History, New, Delete */}
      {state.view === 'chat' && (
        <div className={classes.middleSection}>
          <Tooltip
            content={chat.includeContext ? 'Context: ON — Click to turn OFF' : 'Context: OFF — Click to turn ON'}
            relationship="label"
          >
            <Button
              icon={<DocumentTextRegular style={{ color: chat.includeContext ? '#76B900' : tokens.colorNeutralForeground3 }} />}
              appearance="subtle"
              size="small"
              onClick={() => chat.setIncludeContext(!chat.includeContext)}
            />
          </Tooltip>

          <PromptTemplatesPanel onPick={handlePickTemplate} />

          {onToggleWebSearch && (
            <Tooltip
              content={webSearchEnabled ? 'Web Search: ON — Click to turn OFF' : 'Web Search: OFF — Click to turn ON'}
              relationship="label"
            >
              <Button
                icon={<GlobeRegular style={{ color: webSearchEnabled ? '#76B900' : tokens.colorNeutralForeground3 }} />}
                appearance="subtle"
                size="small"
                onClick={onToggleWebSearch}
              />
            </Tooltip>
          )}

          <Menu>
            <Tooltip content="Chat History" relationship="label">
              <MenuTrigger disableButtonEnhancement>
                <Button icon={<HistoryRegular />} appearance="subtle" size="small" />
              </MenuTrigger>
            </Tooltip>
            <MenuPopover>
              <Input
                className={classes.historySearch}
                placeholder="Search history..."
                value={historyQuery}
                onChange={(_, d) => setHistoryQuery(d.value)}
                contentBefore={<SearchRegular />}
                size="small"
              />
              <MenuList>
                {filteredConversations.length === 0 ? (
                  <MenuItem disabled>
                    {chat.conversations.length === 0 ? 'No chat history' : 'No matches'}
                  </MenuItem>
                ) : (
                  filteredConversations.map(conv => (
                    <MenuItem
                      key={conv.id}
                      onClick={() => chat.setActiveConversation(conv.id)}
                      style={{
                        fontWeight: conv.id === chat.activeConversation?.id ? 600 : 400,
                        backgroundColor: conv.id === chat.activeConversation?.id ? tokens.colorNeutralBackground1Selected : undefined,
                      }}
                    >
                      <div className={classes.historyItem}>
                        <span className={classes.historyItemTitle}>{conv.title}</span>
                        <span className={classes.historyItemMeta}>
                          {conv.messages.length} msgs · {new Date(conv.updatedAt).toLocaleDateString()}
                        </span>
                      </div>
                    </MenuItem>
                  ))
                )}
              </MenuList>
            </MenuPopover>
          </Menu>

          <Tooltip content="New Chat" relationship="label">
            <Button
              icon={<AddRegular />}
              appearance="subtle"
              size="small"
              onClick={chat.createConversation}
            />
          </Tooltip>

          {chat.activeConversation && (
            <Tooltip content="Delete Conversation" relationship="label">
              <Button
                icon={<DeleteRegular />}
                appearance="subtle"
                size="small"
                onClick={() => chat.deleteConversation(chat.activeConversation!.id)}
              />
            </Tooltip>
          )}
        </div>
      )}

      {/* Right: Settings Icon Button */}
      <div className={classes.rightSection}>
        <Tooltip content="Settings" relationship="label">
          <Button
            appearance={state.view === 'settings' ? 'subtle' : 'transparent'}
            icon={<Settings20Regular />}
            onClick={() => dispatch({ type: 'SET_VIEW', payload: 'settings' })}
            size="small"
          />
        </Tooltip>
      </div>
    </div>
  );
};
