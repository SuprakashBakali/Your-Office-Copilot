import React from 'react';
import { makeStyles, tokens } from '@fluentui/react-components';
import { NavigationTabs } from './NavigationTabs';
import { useAppState } from '../../store/AppContext';
import { useChat } from '../../hooks/useChat';
import { ChatPanel } from '../chat/ChatPanel';
import { SettingsPanel } from '../settings/SettingsPanel';
import { NotificationToast } from '../shared/NotificationToast';

const useStyles = makeStyles({
  sidebar: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    width: '100%',
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    overflow: 'hidden',
  },
  content: {
    flexGrow: 1,
    overflowY: 'auto',
    overflowX: 'hidden',
    backgroundColor: tokens.colorNeutralBackground2,
    position: 'relative',
    display: 'flex',
    flexDirection: 'column',
  },
});

export const Sidebar: React.FC = () => {
  const classes = useStyles();
  const { view, host, notification } = useAppState();
  const chat = useChat(host);

  const renderView = () => {
    switch (view) {
      case 'settings':
        return <SettingsPanel />;
      case 'chat':
      default:
        return <ChatPanel chat={chat} />;
    }
  };

  return (
    <div className={classes.sidebar}>
      <NavigationTabs chat={chat} />
      <main className={classes.content}>
        {renderView()}
      </main>
      {notification && (
        <NotificationToast
          type={notification.type}
          message={notification.message}
        />
      )}
    </div>
  );
};
