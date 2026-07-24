import React from 'react';
import { FluentProvider } from '@fluentui/react-components';
import { useTheme } from './hooks/useTheme';
import { AppProvider } from './store/AppContext';
import { Sidebar } from './components/layout/Sidebar';
import './styles/global.css';

export const App: React.FC = () => {
  const { theme } = useTheme();
  return (
    <FluentProvider theme={theme}>
      <AppProvider>
        <Sidebar />
      </AppProvider>
    </FluentProvider>
  );
};
