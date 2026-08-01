import React, { useState } from 'react';
import { makeStyles, tokens, Input, Button, Text, Spinner, Badge } from '@fluentui/react-components';
import { EyeRegular, EyeOffRegular, CheckmarkCircleRegular, ErrorCircleRegular } from '@fluentui/react-icons';
import { useSettings } from '../../hooks/useSettings';
import { useAI } from '../../hooks/useAI';
import { PROVIDER_CONFIGS } from '../../services/ai/ProviderFactory';
import { AIProviderType } from '../../types';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    padding: '10px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: '10px',
    backgroundColor: tokens.colorNeutralBackground1,
    transition: 'border-color 0.2s ease',
    '&:hover': {
      border: `1px solid ${tokens.colorNeutralStroke1}`,
    },
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  providerDot: {
    width: '8px',
    height: '8px',
    borderRadius: '50%',
    flexShrink: 0,
  },
  inputArea: {
    display: 'flex',
    gap: '6px',
    alignItems: 'center',
  },
  input: {
    flexGrow: 1,
  },
  status: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    fontSize: tokens.fontSizeBase200,
    marginTop: '2px',
  },
  success: { color: tokens.colorPaletteGreenForeground1 },
  error: { color: tokens.colorPaletteRedForeground1 },
});

export const ApiKeyManager: React.FC<{ provider: string }> = ({ provider }) => {
  const classes = useStyles();
  const { getApiKey, setApiKey } = useSettings();
  const { testConnection } = useAI();
  const config = PROVIDER_CONFIGS[provider];
  const isOllama = provider === 'ollama';

  const [key, setKey] = useState(getApiKey(provider as AIProviderType) || '');
  const [showKey, setShowKey] = useState(false);
  const [status, setStatus] = useState<'idle' | 'testing' | 'success' | 'error'>('idle');
  const [errorMsg, setErrorMsg] = useState('');

  const hasKey = key.length > 0;

  const handleSave = async () => {
    setApiKey(provider as AIProviderType, key);
    if (key.length > 5) {
      setStatus('testing');
      setErrorMsg('');
      try {
        const result = await testConnection(provider as AIProviderType, key);
        if (result.ok) {
          setStatus('success');
        } else {
          setStatus('error');
          setErrorMsg(result.error || 'Could not verify key — check it and try again');
        }
      } catch (e) {
        setStatus('error');
        setErrorMsg((e as Error).message || 'Unknown error');
      }
    } else if (key.length === 0) {
      setStatus('idle');
    } else {
      setStatus('error');
      setErrorMsg('Key is too short');
    }
  };

  return (
    <div className={classes.container}>
      <div className={classes.header}>
        <span className={classes.providerDot} style={{ backgroundColor: config?.color || '#888' }} />
        <Text weight="semibold" size={300}>{config?.name || provider}</Text>
        {isOllama && <Badge appearance="outline" size="small" color="informative">Local</Badge>}
        {!isOllama && hasKey && status === 'success' && <Badge appearance="filled" size="small" color="success">✓</Badge>}
        {!isOllama && hasKey && status === 'error' && <Badge appearance="filled" size="small" color="danger">✗</Badge>}
      </div>
      {!isOllama ? (
        <>
          <div className={classes.inputArea}>
            <Input
              className={classes.input}
              type={showKey ? 'text' : 'password'}
              value={key}
              onChange={(_, d) => { setKey(d.value); setStatus('idle'); }}
              placeholder={`Enter ${config?.name || provider} API key`}
              size="small"
              contentAfter={
                <Button
                  appearance="transparent"
                  size="small"
                  icon={showKey ? <EyeOffRegular /> : <EyeRegular />}
                  onClick={() => setShowKey(!showKey)}
                />
              }
            />
            <Button
              onClick={handleSave}
              appearance="primary"
              size="small"
              disabled={status === 'testing'}
              icon={status === 'testing' ? <Spinner size="extra-tiny" /> : undefined}
            >
              {status === 'testing' ? 'Testing...' : 'Save'}
            </Button>
          </div>
          {status === 'success' && (
            <div className={`${classes.status} ${classes.success}`}>
              <CheckmarkCircleRegular /> Key verified successfully
            </div>
          )}
          {status === 'error' && (
            <div className={`${classes.status} ${classes.error}`}>
              <ErrorCircleRegular /> {errorMsg || 'Could not verify key — check it and try again'}
            </div>
          )}
        </>
      ) : (
        <div>
          <div className={classes.inputArea}>
            <Input
              className={classes.input}
              value={key || 'http://localhost:11434'}
              onChange={(_, d) => { setKey(d.value); setStatus('idle'); }}
              placeholder="Ollama URL (default: http://localhost:11434)"
              size="small"
            />
            <Button onClick={handleSave} appearance="primary" size="small">Save</Button>
          </div>
          <Text size={100} style={{ color: tokens.colorNeutralForeground3, marginTop: '4px' }}>
            No API key needed — runs locally on your machine
          </Text>
        </div>
      )}
    </div>
  );
};
