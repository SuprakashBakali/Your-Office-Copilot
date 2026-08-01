import React from 'react';
import { makeStyles, tokens } from '@fluentui/react-components';
import { Header } from './Header';
import { NavigationTabs } from './NavigationTabs';
import { useAppState } from '../../store/AppContext';
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
  },
});

export const Sidebar: React.FC = () => {
  const classes = useStyles();
  const { view, notification } = useAppState();

  const renderView = () => {
    switch (view) {
      case 'settings':
        return <SettingsPanel />;
      case 'chat':
      default:
        return <ChatPanel />;
    }
  };

  return (
    <div className={classes.sidebar}>
      <Header />
      <NavigationTabs />
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
