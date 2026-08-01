import React from 'react';
import {
  makeStyles, tokens, Button, Dropdown, Option, Tooltip,
  Menu, MenuItem, MenuTrigger, MenuPopover, MenuList
} from '@fluentui/react-components';
import {
  Chat20Regular, Settings20Regular, HistoryRegular, AddRegular,
  ArrowExportRegular, DeleteRegular, DocumentTextRegular,
  CodeRegular, TextTRegular, WeatherSunny24Regular, WeatherMoon24Regular
} from '@fluentui/react-icons';
import { useAppState, useAppDispatch } from '../../store/AppContext';
import { AppView } from '../../types';
import { UseChatReturn } from '../../hooks/useChat';
import { useTheme } from '../../hooks/useTheme';

const useStyles = makeStyles({
  nav: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '6px',
    padding: '4px 8px',
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
    gap: '4px',
    flexGrow: 1,
    minWidth: 0,
    justifyContent: 'center',
  },
  historyDropdown: {
    flexGrow: 1,
    minWidth: '100px',
    maxWidth: '180px',
  },
  rightSection: {
    display: 'flex',
    alignItems: 'center',
    gap: '2px',
    flexShrink: 0,
  },
});

interface NavigationTabsProps {
  chat: UseChatReturn;
}

export const NavigationTabs: React.FC<NavigationTabsProps> = ({ chat }) => {
  const classes = useStyles();
  const state = useAppState();
  const dispatch = useAppDispatch();
  const { isDark, toggleTheme } = useTheme();

  const handleExport = (format: 'txt' | 'json' | 'markdown') => {
    if (chat.activeConversation) {
      chat.exportConversation(chat.activeConversation.id, format);
    }
  };

  return (
    <div className={classes.nav}>
      {/* Left: Chat Button */}
      <div className={classes.leftSection}>
        <Button
          appearance={state.view === 'chat' ? 'subtle' : 'transparent'}
          icon={<Chat20Regular />}
          onClick={() => dispatch({ type: 'SET_VIEW', payload: 'chat' })}
          style={{ fontWeight: state.view === 'chat' ? 600 : 400 }}
          size="small"
        >
          Chat
        </Button>
      </div>

      {/* Middle: Chat History / New Chat / Export / Delete (when on Chat tab) */}
      {state.view === 'chat' && (
        <div className={classes.middleSection}>
          <Tooltip content="Chat history" relationship="label">
            <HistoryRegular style={{ fontSize: '14px', color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
          </Tooltip>
          <Dropdown
            className={classes.historyDropdown}
            value={chat.activeConversation?.title || 'New Chat'}
            onOptionSelect={(_, data) => chat.setActiveConversation(data.optionValue as string)}
            size="small"
          >
            {chat.conversations.map(conv => (
              <Option key={conv.id} value={conv.id}>{conv.title}</Option>
            ))}
          </Dropdown>

          <Tooltip content="New Chat" relationship="label">
            <Button
              icon={<AddRegular />}
              appearance="subtle"
              size="small"
              onClick={chat.createConversation}
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

          {chat.activeConversation && (
            <Tooltip content="Delete conversation" relationship="label">
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

      {/* Right: Settings Button & Theme Toggle */}
      <div className={classes.rightSection}>
        <Button
          appearance={state.view === 'settings' ? 'subtle' : 'transparent'}
          icon={<Settings20Regular />}
          onClick={() => dispatch({ type: 'SET_VIEW', payload: 'settings' })}
          style={{ fontWeight: state.view === 'settings' ? 600 : 400 }}
          size="small"
        >
          Settings
        </Button>

        <Tooltip content={isDark ? 'Light mode' : 'Dark mode'} relationship="label">
          <Button
            appearance="transparent"
            size="small"
            icon={isDark ? <WeatherSunny24Regular /> : <WeatherMoon24Regular />}
            onClick={toggleTheme}
          />
        </Tooltip>
      </div>
    </div>
  );
};
