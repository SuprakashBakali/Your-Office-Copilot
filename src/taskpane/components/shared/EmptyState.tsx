import React from 'react';
import { makeStyles, tokens, Text, Button } from '@fluentui/react-components';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: '100%',
    padding: '32px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground2,
  },
  icon: {
    fontSize: '48px',
    color: tokens.colorNeutralForeground4,
    marginBottom: '16px',
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
    marginBottom: '8px',
    color: tokens.colorNeutralForeground1,
  },
  description: {
    fontSize: tokens.fontSizeBase300,
    marginBottom: '24px',
  },
});

interface EmptyStateProps {
  icon: React.ReactNode;
  title: string;
  description: string;
  actionText?: string;
  onAction?: () => void;
}

export const EmptyState: React.FC<EmptyStateProps> = ({
  icon,
  title,
  description,
  actionText,
  onAction,
}) => {
  const classes = useStyles();

  return (
    <div className={classes.container}>
      <div className={classes.icon}>{icon}</div>
      <Text className={classes.title}>{title}</Text>
      <Text className={classes.description}>{description}</Text>
      {actionText && onAction && (
        <Button appearance="primary" onClick={onAction}>
          {actionText}
        </Button>
      )}
    </div>
  );
};
