import React, { useState, useEffect } from 'react';
import { makeStyles, tokens, Dropdown, Option, Text, Input } from '@fluentui/react-components';
import { useSettings } from '../../hooks/useSettings';
import { getProvider } from '../../services/ai/ProviderFactory';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  row: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: '12px',
  },
  customInput: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontSize: '11px',
  },
});

export const ModelSelector: React.FC = () => {
  const classes = useStyles();
  const { settings, updateSettings } = useSettings();
  const [customModelId, setCustomModelId] = useState('');
  const [isCustom, setIsCustom] = useState(false);

  let models: Array<{ id: string; name: string }> = [];
  try {
    const provider = getProvider(settings.activeProvider);
    models = provider.getModels();
  } catch {
    models = [];
  }

  // Detect if current model is custom (not in preset list)
  useEffect(() => {
    const presetIds = models.filter(m => m.id !== '__custom__').map(m => m.id);
    if (settings.activeModel && !presetIds.includes(settings.activeModel)) {
      setIsCustom(true);
      setCustomModelId(settings.activeModel);
    } else {
      setIsCustom(false);
    }
  }, [settings.activeProvider, settings.activeModel]);

  const handleSelect = (_: unknown, data: { optionValue?: string }) => {
    if (data.optionValue === '__custom__') {
      setIsCustom(true);
      setCustomModelId('');
    } else {
      setIsCustom(false);
      updateSettings({ activeModel: data.optionValue as string });
    }
  };

  const handleCustomSubmit = () => {
    const trimmed = customModelId.trim();
    if (trimmed) {
      updateSettings({ activeModel: trimmed });
    }
  };

  const displayValue = isCustom
    ? '✏️ Custom Model ID...'
    : models.find(m => m.id === settings.activeModel)?.name || settings.activeModel;

  return (
    <div className={classes.container}>
      <div className={classes.row}>
        <Text>Model</Text>
        <Dropdown
          value={displayValue}
          onOptionSelect={handleSelect}
          size="small"
        >
          {models.map((model) => (
            <Option key={model.id} value={model.id}>{model.name}</Option>
          ))}
        </Dropdown>
      </div>

      {isCustom && (
        <div className={classes.customInput}>
          <Input
            placeholder="e.g. meta/llama-3.1-70b-instruct"
            value={customModelId}
            size="small"
            onChange={(_, data) => setCustomModelId(data.value)}
            onKeyDown={(e) => { if (e.key === 'Enter') handleCustomSubmit(); }}
            onBlur={handleCustomSubmit}
          />
          <Text className={classes.hint}>
            Enter the exact model ID from NVIDIA NIM catalog. Press Enter to apply.
          </Text>
        </div>
      )}
    </div>
  );
};
