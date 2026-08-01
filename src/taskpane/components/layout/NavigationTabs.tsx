import React from 'react';
import {
  makeStyles, tokens, Button, Tooltip,
  Menu, MenuItem, MenuTrigger, MenuPopover, MenuList
} from '@fluentui/react-components';
import {
  Chat20Regular, Settings20Regular, HistoryRegular, AddRegular,
  DeleteRegular
} from '@fluentui/react-icons';
import { useAppState, useAppDispatch } from '../../store/AppContext';
import { UseChatReturn } from '../../hooks/useChat';

const useStyles = makeStyles({
  nav: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '6px 12px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    flexShrink: 0,
    minHeight: '38px',
  },
  leftSection: {
    display: 'flex',
    alignItems: 'center',
    flexShrink: 0,
  },
  middleSection: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    justifyContent: 'center',
    flexGrow: 1,
  },
  rightSection: {
    display: 'flex',
    alignItems: 'center',
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

  return (
    <div className={classes.nav}>
      {/* Left: Chat Icon Button (No Text Label) */}
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

      {/* Middle: History Menu Icon, New Chat Icon, Delete Icon (when in Chat view) */}
      {state.view === 'chat' && (
        <div className={classes.middleSection}>
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Tooltip content="Chat History" relationship="label">
                <Button icon={<HistoryRegular />} appearance="subtle" size="small" />
              </Tooltip>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                {chat.conversations.length === 0 ? (
                  <MenuItem disabled>No chat history</MenuItem>
                ) : (
                  chat.conversations.map(conv => (
                    <MenuItem
                      key={conv.id}
                      onClick={() => chat.setActiveConversation(conv.id)}
                      style={{
                        fontWeight: conv.id === chat.activeConversation?.id ? 600 : 400,
                        backgroundColor: conv.id === chat.activeConversation?.id ? tokens.colorNeutralBackground1Selected : undefined
                      }}
                    >
                      {conv.title}
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

      {/* Right: Settings Icon Button (No Text Label) */}
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
