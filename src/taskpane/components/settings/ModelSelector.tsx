import React from 'react';
import { makeStyles, tokens, Dropdown, Option, Text } from '@fluentui/react-components';
import { useSettings } from '../../hooks/useSettings';

const useStyles = makeStyles({
  row: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: '12px',
  },
  none: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
    fontStyle: 'italic',
  },
});

const PROVIDER_COLORS: Record<string, string> = {
  nvidia: '#76b900', openai: '#10a37f', anthropic: '#d97706',
  gemini: '#4285f4', groq: '#f55036', openrouter: '#8b5cf6', ollama: '#6b7280',
};

export const ModelSelector: React.FC = () => {
  const classes = useStyles();
  const { settings, updateSettings } = useSettings();
  const models = settings.customModels || [];
  const activeId = settings.activeCustomModelId || '';
  const activeModel = models.find(m => m.id === activeId);

  if (models.length === 0) {
    return (
      <div className={classes.row}>
        <Text>Model</Text>
        <Text className={classes.none}>No models — add one in Settings ↓</Text>
      </div>
    );
  }

  return (
    <div className={classes.row}>
      <Text>Model</Text>
      <Dropdown
        value={activeModel?.name || 'Select a model'}
        onOptionSelect={(_, d) => {
          const chosen = models.find(m => m.id === d.optionValue);
          if (chosen) {
            updateSettings({
              activeCustomModelId: chosen.id,
              activeProvider: chosen.provider,
              activeModel: chosen.modelId,
            });
          }
        }}
        size="small"
      >
        {models.map(m => (
          <Option key={m.id} value={m.id} text={m.name}>
            <span style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <span style={{
                width: 6, height: 6, borderRadius: '50%',
                backgroundColor: PROVIDER_COLORS[m.provider] || '#888',
                display: 'inline-block', flexShrink: 0,
              }} />
              {m.name}
            </span>
          </Option>
        ))}
      </Dropdown>
    </div>
  );
};
