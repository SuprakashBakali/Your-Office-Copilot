import React from 'react';
import { makeStyles, Dropdown, Option, Text } from '@fluentui/react-components';
import { useSettings } from '../../hooks/useSettings';
import { getProvider } from '../../services/ai/ProviderFactory';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: '12px',
  }
});

export const ModelSelector: React.FC = () => {
  const classes = useStyles();
  const { settings, updateSettings } = useSettings();

  let models: Array<{ id: string; name: string }> = [];
  try {
    const provider = getProvider(settings.activeProvider);
    models = provider.getModels();
  } catch {
    models = [];
  }

  return (
    <div className={classes.container}>
      <Text>Model</Text>
      <Dropdown
        value={models.find(m => m.id === settings.activeModel)?.name || settings.activeModel}
        onOptionSelect={(_, data) => updateSettings({ activeModel: data.optionValue as string })}
      >
        {models.map((model) => (
          <Option key={model.id} value={model.id}>{model.name}</Option>
        ))}
      </Dropdown>
    </div>
  );
};
