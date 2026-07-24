import React from 'react';
import { Button, makeStyles, tokens, Text, Tooltip, Badge } from '@fluentui/react-components';
import { Settings24Regular, WeatherSunny24Regular, WeatherMoon24Regular, Bot24Regular } from '@fluentui/react-icons';
import { useAppState, useAppDispatch } from '../../store/AppContext';
import { useTheme } from '../../hooks/useTheme';
import { loadSettings } from '../../utils/storage';
import { PROVIDER_CONFIGS } from '../../services/ai/ProviderFactory';

const useStyles = makeStyles({
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingLeft: '12px',
    paddingRight: '8px',
    paddingTop: '8px',
    paddingBottom: '8px',
    minHeight: '48px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
  titleArea: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
  },
  iconWrapper: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '34px',
    height: '34px',
    borderRadius: '10px',
    backgroundImage: 'linear-gradient(135deg, #76B900 0%, #4a8c00 50%, #2A4E00 100%)',
    color: '#ffffff',
    boxShadow: '0 2px 8px rgba(118, 185, 0, 0.3)',
    transition: 'transform 0.2s ease, box-shadow 0.2s ease',
    ':hover': {
      transform: 'scale(1.05)',
      boxShadow: '0 4px 12px rgba(118, 185, 0, 0.5)',
    },
  },
  titleText: {
    display: 'flex',
    flexDirection: 'column',
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    color: tokens.colorNeutralForeground1,
  },
  subtitle: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    lineHeight: tokens.lineHeightBase100,
  },
  providerDot: {
    width: '6px',
    height: '6px',
    borderRadius: '50%',
    display: 'inline-block',
  },
  actions: {
    display: 'flex',
    gap: '2px',
    alignItems: 'center',
  },
});

export const Header: React.FC = () => {
  const classes = useStyles();
  const state = useAppState();
  const dispatch = useAppDispatch();
  const { isDark, toggleTheme } = useTheme();
  const settings = loadSettings();
  const providerConfig = PROVIDER_CONFIGS[settings.activeProvider];

  const hostLabel = state.host === 'Unknown' ? 'Excel' : state.host;

  return (
    <header className={classes.header}>
      <div className={classes.titleArea}>
        <div className={classes.iconWrapper}>
          <Bot24Regular />
        </div>
        <div className={classes.titleText}>
          <Text className={classes.title}>Your Co-Pilot</Text>
          <span className={classes.subtitle}>
            <span className={classes.providerDot} style={{ backgroundColor: providerConfig?.color || '#76B900' }} />
            {providerConfig?.name || 'NVIDIA NIM'} · {hostLabel}
          </span>
        </div>
      </div>
      <div className={classes.actions}>
        <Tooltip content={isDark ? 'Light mode' : 'Dark mode'} relationship="label">
          <Button
            appearance="transparent"
            size="small"
            icon={isDark ? <WeatherSunny24Regular /> : <WeatherMoon24Regular />}
            onClick={toggleTheme}
          />
        </Tooltip>
        <Tooltip content="Settings" relationship="label">
          <Button
            appearance="transparent"
            size="small"
            icon={<Settings24Regular />}
            onClick={() => dispatch({ type: 'SET_VIEW', payload: 'settings' })}
          />
        </Tooltip>
      </div>
    </header>
  );
};
