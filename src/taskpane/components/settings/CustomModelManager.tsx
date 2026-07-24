import React, { useState } from 'react';
import {
  makeStyles, tokens, Button, Input, Dropdown, Option, Text,
  Badge, Spinner, Tooltip,
} from '@fluentui/react-components';
import {
  AddRegular, DeleteRegular, CheckmarkCircleRegular,
  ErrorCircleRegular, EyeRegular, EyeOffRegular, RadioButtonRegular, RecordRegular,
  PlayRegular, TimerRegular, TrophyRegular,
} from '@fluentui/react-icons';
import { useSettings } from '../../hooks/useSettings';
import { AIProviderType, CustomModel } from '../../types';
import { GenericOpenAIProvider } from '../../services/ai/GenericOpenAIProvider';
import { getProvider } from '../../services/ai/ProviderFactory';

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
  leaderboard: {
    marginTop: '16px',
    padding: '16px',
    borderRadius: '10px',
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  leaderboardRow: {
    display: 'flex',
    justifyContent: 'space-between',
    padding: '8px 0',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    '&:last-child': { borderBottom: 'none' },
  },
  leaderboardRank: {
    width: '24px',
    fontWeight: 'bold',
    color: tokens.colorBrandForeground1,
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

const ModelCard: React.FC<ModelCardProps> = ({ model, isActive, onSetActive, onUpdate, onDelete }) => {
  const classes = useStyles();
  const [showKey, setShowKey] = useState(false);
  const [testStatus, setTestStatus] = useState<'idle' | 'testing' | 'ok' | 'error'>('idle');
  const [testError, setTestError] = useState('');

  const update = (patch: Partial<CustomModel>) => onUpdate({ ...model, ...patch });

  const handleTest = async () => {
    if (!model.apiKey || !model.modelId || !model.baseUrl) return;
    setTestStatus('testing');
    setTestError('');
    try {
      const provider = new GenericOpenAIProvider(model.baseUrl, model.apiKey);
      const response = await provider.chat({
        model: model.modelId,
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

  const providerColor = PROVIDERS.find(p => p.id === model.provider)?.color || '#888';

  return (
    <div className={`${classes.card} ${isActive ? classes.activeCard : ''}`}>
      {/* Active selector + delete */}
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
          value={model.name}
          placeholder="Model display name"
          size="small"
          onChange={(_, d) => update({ name: d.value })}
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
          value={PROVIDERS.find(p => p.id === model.provider)?.label || model.provider}
          onOptionSelect={(_, d) => {
            const chosen = PROVIDERS.find(p => p.id === d.optionValue);
            update({
              provider: d.optionValue as AIProviderType,
              // Auto-fill base URL when provider changes
              baseUrl: model.baseUrl || chosen?.baseUrl || '',
            });
          }}
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
          value={model.modelId}
          placeholder="e.g. meta/llama-3.1-70b-instruct"
          size="small"
          onChange={(_, d) => update({ modelId: d.value })}
        />
      </div>

      {/* Base URL */}
      <div className={classes.row}>
        <Text className={classes.label}>Base URL</Text>
        <Input
          className={classes.flex1}
          value={model.baseUrl || ''}
          placeholder={PROVIDERS.find(p => p.id === model.provider)?.baseUrl || 'https://...'}
          size="small"
          onChange={(_, d) => update({ baseUrl: d.value })}
        />
      </div>

      {/* API Key */}
      {model.provider !== 'ollama' && (
        <div className={classes.row}>
          <Text className={classes.label}>API Key</Text>
          <Input
            className={classes.flex1}
            type={showKey ? 'text' : 'password'}
            value={model.apiKey}
            placeholder="Paste API key here"
            size="small"
            onChange={(_, d) => { update({ apiKey: d.value }); setTestStatus('idle'); }}
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
            disabled={testStatus === 'testing' || !model.apiKey || !model.modelId || !model.baseUrl}
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
  const [testRankings, setTestRankings] = useState<{ id: string, name: string, latency: number, status: 'ok'|'error', error?: string }[]>([]);
  const [isTestingAll, setIsTestingAll] = useState(false);

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
    saveModels(updated);
    // Auto-set as active if it's the first
    if (updated.length === 1) {
      updateSettings({ customModels: updated, activeCustomModelId: newModel.id });
    }
  };

  const updateModel = (updated: CustomModel) => {
    saveModels(models.map(m => m.id === updated.id ? updated : m));
  };

  const deleteModel = (id: string) => {
    const updated = models.filter(m => m.id !== id);
    saveModels(updated);
    if (activeId === id) {
      updateSettings({ customModels: updated, activeCustomModelId: updated[0]?.id || '' });
    }
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

  const handleTestAll = async () => {
    setIsTestingAll(true);
    setTestRankings([]);
    
    const results = await Promise.all(models.map(async (model) => {
      if (!model.apiKey || !model.modelId) {
        return { id: model.id, name: model.name, latency: 999999, status: 'error' as const, error: 'Missing API Key or Model ID' };
      }
      const start = performance.now();
      try {
        const provider = model.baseUrl ? new GenericOpenAIProvider(model.baseUrl, model.apiKey) : getProvider(model.provider);
        if (!model.baseUrl) provider.setApiKey(model.apiKey);
        
        await provider.chat({
          model: model.modelId,
          messages: [{ role: 'user', content: 'Hi' }],
          maxTokens: 5,
          stream: false,
        });
        const end = performance.now();
        return { id: model.id, name: model.name, latency: Math.round(end - start), status: 'ok' as const };
      } catch (e) {
        const end = performance.now();
        return { id: model.id, name: model.name, latency: Math.round(end - start), status: 'error' as const, error: (e as Error).message };
      }
    }));
    
    setTestRankings(results.sort((a, b) => a.latency - b.latency));
    setIsTestingAll(false);
  };

  return (
    <div className={classes.root}>
      <div className={classes.header}>
        <Text weight="semibold">My Models ({models.length})</Text>
        <div style={{ display: 'flex', gap: '8px' }}>
          <Button icon={isTestingAll ? <Spinner size="extra-tiny" /> : <PlayRegular />} size="small" appearance="secondary" onClick={handleTestAll} disabled={isTestingAll || models.length === 0}>
            {isTestingAll ? 'Testing...' : 'Test All'}
          </Button>
          <Button icon={<AddRegular />} size="small" appearance="primary" onClick={addModel}>
            Add Model
          </Button>
        </div>
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

      {testRankings.length > 0 && (
        <div className={classes.leaderboard}>
          <Text weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '8px' }}>
            <TrophyRegular style={{ color: '#d97706' }}/> API Latency Leaderboard
          </Text>
          {testRankings.map((r, i) => (
            <div key={r.id} className={classes.leaderboardRow}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Text className={classes.leaderboardRank}>#{i + 1}</Text>
                <Text>{r.name}</Text>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                {r.status === 'ok' ? (
                  <Text style={{ color: tokens.colorPaletteGreenForeground1, display: 'flex', alignItems: 'center', gap: '4px' }}>
                    <TimerRegular /> {r.latency}ms
                  </Text>
                ) : (
                  <Text style={{ color: tokens.colorPaletteRedForeground1, fontSize: '11px', maxWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    Error: {r.error}
                  </Text>
                )}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
