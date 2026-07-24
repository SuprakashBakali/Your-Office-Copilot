import React from 'react';
import { makeStyles, tokens, Text } from '@fluentui/react-components';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  dots: {
    display: 'flex',
    gap: '4px',
  },
  dot: {
    width: '6px',
    height: '6px',
    borderRadius: '50%',
    backgroundColor: tokens.colorBrandBackground,
    animationName: {
      '0%': { transform: 'translateY(0)' },
      '40%': { transform: 'translateY(-6px)' },
      '80%': { transform: 'translateY(0)' },
      '100%': { transform: 'translateY(0)' },
    },
    animationDuration: '1.4s',
    animationIterationCount: 'infinite',
    animationTimingFunction: 'ease-in-out',
  },
  dot1: { animationDelay: '-0.32s' },
  dot2: { animationDelay: '-0.16s' },
  dot3: { animationDelay: '0s' },
});

interface LoadingDotsProps {
  label?: string;
}

export const LoadingDots: React.FC<LoadingDotsProps> = ({ label }) => {
  const classes = useStyles();

  return (
    <div className={classes.container}>
      <div className={classes.dots}>
        <div className={`${classes.dot} ${classes.dot1}`} />
        <div className={`${classes.dot} ${classes.dot2}`} />
        <div className={`${classes.dot} ${classes.dot3}`} />
      </div>
      {label && <Text size={200}>{label}</Text>}
    </div>
  );
};
