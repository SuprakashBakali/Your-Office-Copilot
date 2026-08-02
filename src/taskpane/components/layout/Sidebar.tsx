import React, { useState } from 'react';
import { makeStyles, tokens, Spinner, Text } from '@fluentui/react-components';
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
  loading: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '12px',
    height: '100%',
    color: tokens.colorNeutralForeground3,
  },
});

export const Sidebar: React.FC = () => {
  const classes = useStyles();
  const { view, host, notification } = useAppState();
  // useChat is called with the resolved host — but only after Office.js
  // detects the host. Before that, host is 'Unknown'. We still call useChat
  // unconditionally (hooks rule), but we DON'T render ChatPanel until the
  // host resolves, so no conversations are created with hostApp='Unknown'
  // (which would lead to wrong system prompts and context).
  const chat = useChat(host);
  // Web search toggle lives here so it persists across chat re-mounts and
  // stays in sync with the nav bar button.
  const [webSearchEnabled, setWebSearchEnabled] = useState(false);

  const renderView = () => {
    switch (view) {
      case 'settings':
        return <SettingsPanel />;
      case 'chat':
      default:
        // Defer rendering ChatPanel until the Office host is resolved.
        // This prevents conversations from being created with hostApp='Unknown'.
        if (host === 'Unknown') {
          return (
            <div className={classes.loading}>
              <Spinner size="medium" />
              <Text size={200}>Connecting to Office…</Text>
            </div>
          );
        }
        return <ChatPanel chat={chat} webSearchEnabled={webSearchEnabled} />;
    }
  };

  return (
    <div className={classes.sidebar}>
      <NavigationTabs
        chat={chat}
        webSearchEnabled={webSearchEnabled}
        onToggleWebSearch={() => setWebSearchEnabled(v => !v)}
      />
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
