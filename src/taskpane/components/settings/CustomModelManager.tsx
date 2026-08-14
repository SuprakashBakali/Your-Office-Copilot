import React, { useState, useRef } from 'react';
import {
  makeStyles, tokens, Button, Input, Dropdown, Option, Text,
  Badge, Tooltip, Spinner
} from '@fluentui/react-components';
import {
  AddRegular, DeleteRegular,
  EyeRegular, EyeOffRegular, RadioButtonRegular, RecordRegular,
  CheckmarkCircleRegular, ErrorCircleRegular
} from '@fluentui/react-icons';
import { useSettings } from '../../hooks/useSettings';
import { AIProviderType, CustomModel } from '../../types';
import { GenericOpenAIProvider } from '../../services/ai/GenericOpenAIProvider';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  card: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: '10px',
    padding: '12px',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    backgroundColor: tokens.colorNeutralBackground1,
    transition: 'border-color 0.2s',
    '&:hover': { border: `1px solid ${tokens.colorNeutralStroke1}` },
  },
  activeCard: {
    border: `1px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorBrandBackground2,
  },
  row: {
    display: 'flex',
    gap: '6px',
    alignItems: 'center',
  },
  label: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    minWidth: '58px',
  },
  flex1: { flex: 1 },
  status: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    fontSize: tokens.fontSizeBase100,
    marginTop: '2px',
  },
  success: { color: tokens.colorPaletteGreenForeground1 },
  error: { color: tokens.colorPaletteRedForeground1 },
  empty: {
    textAlign: 'center',
    padding: '24px',
    color: tokens.colorNeutralForeground3,
    border: `1px dashed ${tokens.colorNeutralStroke2}`,
    borderRadius: '10px',
  },
});

const PROVIDERS: { id: AIProviderType; label: string; color: string; baseUrl: string }[] = [
  { id: 'nvidia',      label: 'NVIDIA NIM',    color: '#76b900', baseUrl: 'https://integrate.api.nvidia.com/v1' },
  { id: 'openai',      label: 'OpenAI',         color: '#10a37f', baseUrl: 'https://api.openai.com/v1' },
  { id: 'anthropic',   label: 'Anthropic',      color: '#d97706', baseUrl: 'https://api.anthropic.com/v1' },
  { id: 'gemini',      label: 'Google Gemini',  color: '#4285f4', baseUrl: 'https://generativelanguage.googleapis.com/v1beta/openai' },
  { id: 'groq',        label: 'Groq',           color: '#f55036', baseUrl: 'https://api.groq.com/openai/v1' },
  { id: 'openrouter',  label: 'OpenRouter',     color: '#8b5cf6', baseUrl: 'https://openrouter.ai/api/v1' },
  { id: 'ollama',      label: 'Ollama (Local)', color: '#6b7280', baseUrl: 'http://localhost:11434/v1' },
];

function generateId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

interface ModelCardProps {
  model: CustomModel;
  isActive: boolean;
  onSetActive: () => void;
  onUpdate: (updated: CustomModel) => void;
  onDelete: () => void;
}

/**
 * ModelCard uses local state for all text inputs so keystrokes are never
 * lost due to parent re-renders. It only flushes to the parent store when
 * the input loses focus (onBlur) or when a non-text field changes (dropdown).
 */
const ModelCard: React.FC<ModelCardProps> = ({ model, isActive, onSetActive, onUpdate, onDelete }) => {
  const classes = useStyles();
  const [showKey, setShowKey] = useState(false);
  const [testStatus, setTestStatus] = useState<'idle' | 'testing' | 'ok' | 'error'>('idle');
  const [testError, setTestError] = useState('');

  // Local draft state — mirrors the model so inputs are always controlled,
  // but we only call onUpdate (which saves to settings) on blur / dropdown change.
  const [draft, setDraft] = useState<CustomModel>(model);
  // Keep draft in sync if the parent model reference changes externally
  //  (e.g. after setting it as active), but NOT while the user is typing.
  const isTypingRef = useRef(false);
  React.useEffect(() => {
    if (!isTypingRef.current) {
      setDraft(model);
    }
  }, [model]);

  const handleTextChange = (field: keyof CustomModel, value: string) => {
    isTypingRef.current = true;
    setDraft(prev => ({ ...prev, [field]: value }));
    if (field === 'apiKey' || field === 'modelId' || field === 'baseUrl') {
      setTestStatus('idle');
    }
  };

  const handleBlur = () => {
    isTypingRef.current = false;
    // Only save if something actually changed
    if (
      draft.name !== model.name ||
      draft.modelId !== model.modelId ||
      draft.baseUrl !== model.baseUrl ||
      draft.apiKey !== model.apiKey
    ) {
      onUpdate({ ...model, ...draft });
    }
  };

  const handleDropdownChange = (provider: AIProviderType) => {
    const chosen = PROVIDERS.find(p => p.id === provider);
    const updated: CustomModel = {
      ...draft,
      provider,
      baseUrl: draft.baseUrl || chosen?.baseUrl || '',
    };
    setDraft(updated);
    onUpdate(updated);
    setTestStatus('idle');
  };

  const handleTest = async () => {
    if (!draft.apiKey || !draft.modelId || !draft.baseUrl) return;
    setTestStatus('testing');
    setTestError('');
    // Flush draft to parent before testing
    onUpdate({ ...model, ...draft });
    try {
      const provider = new GenericOpenAIProvider(draft.baseUrl, draft.apiKey);
      const response = await provider.chat({
        model: draft.modelId,
        messages: [{ role: 'user', content: 'Hi' }],
        maxTokens: 5,
        stream: false,
      });
      setTestStatus(response.content ? 'ok' : 'error');
    } catch (e) {
      setTestStatus('error');
      setTestError((e as Error).message);
    }
  };

  return (
    <div className={`${classes.card} ${isActive ? classes.activeCard : ''}`}>
      {/* Active selector + name + delete */}
      <div className={classes.row}>
        <Tooltip content={isActive ? 'Active model' : 'Set as active'} relationship="label">
          <Button
            appearance="transparent"
            size="small"
            icon={isActive ? <RecordRegular style={{ color: tokens.colorBrandForeground1 }} /> : <RadioButtonRegular />}
            onClick={onSetActive}
          />
        </Tooltip>
        <Input
          className={classes.flex1}
          value={draft.name}
          placeholder="Model display name"
          size="small"
          onChange={(_, d) => handleTextChange('name', d.value)}
          onBlur={handleBlur}
        />
        {isActive && <Badge appearance="filled" color="brand" size="small">Active</Badge>}
        <Button
          appearance="subtle"
          size="small"
          icon={<DeleteRegular />}
          onClick={onDelete}
          style={{ color: tokens.colorPaletteRedForeground1 }}
        />
      </div>

      {/* Provider */}
      <div className={classes.row}>
        <Text className={classes.label}>Provider</Text>
        <Dropdown
          value={PROVIDERS.find(p => p.id === draft.provider)?.label || draft.provider}
          onOptionSelect={(_, d) => handleDropdownChange(d.optionValue as AIProviderType)}
          size="small"
          style={{ flex: 1 }}
        >
          {PROVIDERS.map(p => (
            <Option key={p.id} value={p.id} text={p.label}>
              <span style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span style={{ width: 8, height: 8, borderRadius: '50%', backgroundColor: p.color, display: 'inline-block' }} />
                {p.label}
              </span>
            </Option>
          ))}
        </Dropdown>
      </div>

      {/* Model ID */}
      <div className={classes.row}>
        <Text className={classes.label}>Model ID</Text>
        <Input
          className={classes.flex1}
          value={draft.modelId}
          placeholder="e.g. meta/llama-3.1-70b-instruct"
          size="small"
          onChange={(_, d) => handleTextChange('modelId', d.value)}
          onBlur={handleBlur}
        />
      </div>

      {/* Base URL */}
      <div className={classes.row}>
        <Text className={classes.label}>Base URL</Text>
        <Input
          className={classes.flex1}
          value={draft.baseUrl || ''}
          placeholder={PROVIDERS.find(p => p.id === draft.provider)?.baseUrl || 'https://...'}
          size="small"
          onChange={(_, d) => handleTextChange('baseUrl', d.value)}
          onBlur={handleBlur}
        />
      </div>

      {/* API Key */}
      {draft.provider !== 'ollama' && (
        <div className={classes.row}>
          <Text className={classes.label}>API Key</Text>
          <Input
            className={classes.flex1}
            type={showKey ? 'text' : 'password'}
            value={draft.apiKey}
            placeholder="Paste API key here"
            size="small"
            onChange={(_, d) => handleTextChange('apiKey', d.value)}
            onBlur={handleBlur}
            contentAfter={
              <Button
                appearance="transparent"
                size="small"
                icon={showKey ? <EyeOffRegular /> : <EyeRegular />}
                onClick={() => setShowKey(s => !s)}
              />
            }
          />
          <Button
            size="small"
            appearance="outline"
            onClick={handleTest}
            disabled={testStatus === 'testing' || !draft.apiKey || !draft.modelId || !draft.baseUrl}
            icon={testStatus === 'testing' ? <Spinner size="extra-tiny" /> : undefined}
          >
            {testStatus === 'testing' ? '' : 'Test'}
          </Button>
        </div>
      )}

      {/* Test status */}
      {testStatus === 'ok' && (
        <div className={`${classes.status} ${classes.success}`}>
          <CheckmarkCircleRegular /> Connected successfully
        </div>
      )}
      {testStatus === 'error' && (
        <div className={`${classes.status} ${classes.error}`}>
          <ErrorCircleRegular /> {testError || 'Connection failed — check model ID and API key'}
        </div>
      )}
    </div>
  );
};

export const CustomModelManager: React.FC = () => {
  const classes = useStyles();
  const { settings, updateSettings } = useSettings();
  const models = settings.customModels || [];
  const activeId = settings.activeCustomModelId || '';

  const saveModels = (updated: CustomModel[]) => updateSettings({ customModels: updated });

  const addModel = () => {
    const newModel: CustomModel = {
      id: generateId(),
      name: 'New Model',
      modelId: '',
      provider: 'nvidia',
      apiKey: '',
    };
    const updated = [...models, newModel];
    // Auto-set as active if it's the first
    updateSettings({
      customModels: updated,
      ...(updated.length === 1 ? { activeCustomModelId: newModel.id } : {}),
    });
  };

  const updateModel = (updated: CustomModel) => {
    saveModels(models.map(m => m.id === updated.id ? updated : m));
  };

  const deleteModel = (id: string) => {
    const updated = models.filter(m => m.id !== id);
    updateSettings({
      customModels: updated,
      ...(activeId === id ? { activeCustomModelId: updated[0]?.id || '' } : {}),
    });
  };

  const setActive = (id: string) => {
    const model = models.find(m => m.id === id);
    if (model) {
      updateSettings({
        activeCustomModelId: id,
        activeProvider: model.provider,
        activeModel: model.modelId,
      });
    }
  };

  return (
    <div className={classes.root}>
      <div className={classes.header}>
        <Text weight="semibold">My Models ({models.length})</Text>
        <Button icon={<AddRegular />} size="small" appearance="primary" onClick={addModel}>
          Add Model
        </Button>
      </div>

      {models.length === 0 ? (
        <div className={classes.empty}>
          <Text size={200}>No models added yet.<br />Click <strong>Add Model</strong> to get started.</Text>
        </div>
      ) : (
        models.map(model => (
          <ModelCard
            key={model.id}
            model={model}
            isActive={activeId === model.id}
            onSetActive={() => setActive(model.id)}
            onUpdate={updateModel}
            onDelete={() => deleteModel(model.id)}
          />
        ))
      )}
    </div>
  );
};
