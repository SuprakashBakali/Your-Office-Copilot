import { useState, useEffect, useCallback } from 'react';
import { AppSettings, AIProviderType } from '../types';
import { loadSettings, saveSettings, DEFAULT_SETTINGS } from '../utils/storage';

export function useSettings() {
  const [settings, setSettingsState] = useState<AppSettings>(DEFAULT_SETTINGS);

  useEffect(() => {
    setSettingsState(loadSettings());
  }, []);

  const updateSettings = useCallback((newSettings: Partial<AppSettings>) => {
    const current = loadSettings();
    const updated = { ...current, ...newSettings };
    saveSettings(updated);
    setSettingsState(updated);
  }, []);

  const resetSettings = useCallback(() => {
    saveSettings(DEFAULT_SETTINGS);
    setSettingsState(DEFAULT_SETTINGS);
  }, []);

  const getApiKey = useCallback((provider?: AIProviderType) => {
    const p = provider || settings.activeProvider;
    return settings.apiKeys[p] || '';
  }, [settings]);

  const setApiKey = useCallback((provider: AIProviderType, apiKey: string) => {
    const current = loadSettings();
    const updatedKeys = { ...current.apiKeys, [provider]: apiKey };
    const updated = { ...current, apiKeys: updatedKeys };
    saveSettings(updated);
    setSettingsState(updated);
  }, []);

  return {
    settings,
    updateSettings,
    resetSettings,
    getApiKey,
    setApiKey,
  };
}
