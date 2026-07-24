import React from 'react';
import { makeStyles, tokens } from '@fluentui/react-components';
import { Header } from './Header';
import { NavigationTabs } from './NavigationTabs';
import { useAppState } from '../../store/AppContext';
import { ChatPanel } from '../chat/ChatPanel';
import { PromptTemplates } from '../chat/PromptTemplates';
import { SettingsPanel } from '../settings/SettingsPanel';
import { ExcelPanel } from '../excel/ExcelPanel';
import { WordPanel } from '../word/WordPanel';
import { PowerPointPanel } from '../powerpoint/PowerPointPanel';
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
  const { view, host, notification } = useAppState();

  const renderToolsPanel = () => {
    switch (host) {
      case 'Word':
        return <WordPanel />;
      case 'PowerPoint':
        return <PowerPointPanel />;
      case 'Excel':
      default:
        return <ExcelPanel />;
    }
  };

  const renderView = () => {
    switch (view) {
      case 'chat':
        return <ChatPanel />;
      case 'tools':
        return renderToolsPanel();
      case 'templates':
        return <PromptTemplates />;
      case 'settings':
        return <SettingsPanel />;
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
