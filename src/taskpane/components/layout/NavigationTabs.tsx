import React from 'react';
import { TabList, Tab, makeStyles, tokens } from '@fluentui/react-components';
import { Chat20Regular, Settings20Regular } from '@fluentui/react-icons';
import { useAppState, useAppDispatch } from '../../store/AppContext';
import { AppView } from '../../types';

const useStyles = makeStyles({
  nav: {
    padding: '2px 8px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    flexShrink: 0,
  },
});

export const NavigationTabs: React.FC = () => {
  const classes = useStyles();
  const state = useAppState();
  const dispatch = useAppDispatch();

  const handleTabSelect = (_event: any, data: { value: unknown }) => {
    dispatch({ type: 'SET_VIEW', payload: data.value as AppView });
  };

  return (
    <div className={classes.nav}>
      <TabList
        selectedValue={state.view}
        onTabSelect={handleTabSelect}
        size="small"
        appearance="subtle"
      >
        <Tab value="chat" icon={<Chat20Regular />}>Chat</Tab>
        <Tab value="settings" icon={<Settings20Regular />}>Settings</Tab>
      </TabList>
    </div>
  );
};
