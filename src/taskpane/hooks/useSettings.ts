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
    // Use the React state value as the base instead of re-reading from disk.
    // loadSettings() does a synchronous localStorage.getItem + JSON.parse on
    // every call — if multiple components call updateSettings in rapid
    // succession (e.g. slider dragging), each reads stale state and
    // overwrites the previous call's changes.
    setSettingsState(prev => {
      const updated = { ...prev, ...newSettings };
      saveSettings(updated);
      return updated;
    });
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
    // Use functional state update to avoid reading stale settings.
    setSettingsState(prev => {
      const updatedKeys = { ...prev.apiKeys, [provider]: apiKey };
      const updated = { ...prev, apiKeys: updatedKeys };
      saveSettings(updated);
      return updated;
    });
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
