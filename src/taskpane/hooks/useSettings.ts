import { useState, useEffect, useCallback } from 'react';
import { AppSettings, AIProviderType } from '../types';
import { loadSettings, saveSettings, DEFAULT_SETTINGS } from '../utils/storage';

export function useSettings() {
  const [settings, setSettingsState] = useState<AppSettings>(DEFAULT_SETTINGS);

  useEffect(() => {
    setSettingsState(loadSettings());

    const handleSettingsUpdate = () => {
      setSettingsState(loadSettings());
    };
    window.addEventListener('settingsUpdated', handleSettingsUpdate);
    return () => window.removeEventListener('settingsUpdated', handleSettingsUpdate);
  }, []);

  const updateSettings = useCallback((newSettings: Partial<AppSettings>) => {
    const current = loadSettings();
    const updated = { ...current, ...newSettings };
    saveSettings(updated);
    setSettingsState(updated);
    window.dispatchEvent(new Event('settingsUpdated'));
  }, []);

  const resetSettings = useCallback(() => {
    saveSettings(DEFAULT_SETTINGS);
    setSettingsState(DEFAULT_SETTINGS);
    window.dispatchEvent(new Event('settingsUpdated'));
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
    window.dispatchEvent(new Event('settingsUpdated'));
  }, []);

  return {
    settings,
    updateSettings,
    resetSettings,
    getApiKey,
    setApiKey,
  };
}
