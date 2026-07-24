import React from 'react';
import { makeStyles, tokens, Text } from '@fluentui/react-components';

const useStyles = makeStyles({
  container: {
    position: 'fixed',
    bottom: '16px',
    left: '16px',
    right: '16px',
    zIndex: 1000,
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 14px',
    borderRadius: '10px',
    boxShadow: `0 4px 16px ${tokens.colorNeutralShadowAmbient}`,
    animationName: {
      '0%': { opacity: 0, transform: 'translateY(10px)' },
      '100%': { opacity: 1, transform: 'translateY(0)' },
    },
    animationDuration: '0.3s',
    animationTimingFunction: 'ease-out',
    animationFillMode: 'forwards',
  },
  success: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    borderLeft: `4px solid ${tokens.colorPaletteGreenBorder2}`,
    color: tokens.colorPaletteGreenForeground1,
  },
  error: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    borderLeft: `4px solid ${tokens.colorPaletteRedBorder2}`,
    color: tokens.colorPaletteRedForeground1,
  },
  warning: {
    backgroundColor: tokens.colorPaletteYellowBackground1,
    borderLeft: `4px solid ${tokens.colorPaletteYellowBorder2}`,
    color: tokens.colorPaletteYellowForeground1,
  },
  info: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderLeft: `4px solid ${tokens.colorBrandStroke1}`,
    color: tokens.colorNeutralForeground1,
  },
  icon: {
    fontSize: '16px',
    flexShrink: 0,
  },
});

const ICONS: Record<string, string> = {
  success: '✅',
  error: '❌',
  warning: '⚠️',
  info: 'ℹ️',
};

interface NotificationToastProps {
  type: 'success' | 'error' | 'warning' | 'info';
  message: string;
}

export const NotificationToast: React.FC<NotificationToastProps> = ({ type, message }) => {
  const classes = useStyles();

  return (
    <div className={`${classes.container} ${classes[type]}`}>
      <span className={classes.icon}>{ICONS[type]}</span>
      <Text size={200} weight="semibold">{message}</Text>
    </div>
  );
};
