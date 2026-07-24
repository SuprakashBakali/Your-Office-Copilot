import { useState, useEffect, useCallback } from 'react';
import { Theme } from '@fluentui/react-components';
import { ThemeMode } from '../types';
import { loadSettings, saveSettings } from '../utils/storage';
import { customLightTheme, customDarkTheme } from '../styles/theme';

export function useTheme() {
  const [mode, setMode] = useState<ThemeMode>('auto');
  const [isDark, setIsDark] = useState<boolean>(false);

  useEffect(() => {
    const settings = loadSettings();
    if (settings.theme) {
      setMode(settings.theme);
    }
  }, []);

  useEffect(() => {
    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');

    const updateTheme = () => {
      if (mode === 'auto') {
        setIsDark(mediaQuery.matches);
      } else {
        setIsDark(mode === 'dark');
      }
    };

    updateTheme();

    const listener = (e: MediaQueryListEvent) => {
      if (mode === 'auto') {
        setIsDark(e.matches);
      }
    };

    mediaQuery.addEventListener('change', listener);
    return () => mediaQuery.removeEventListener('change', listener);
  }, [mode]);

  // Apply body class for CSS hooks
  useEffect(() => {
    document.body.classList.toggle('dark-theme', isDark);
    document.body.classList.toggle('light-theme', !isDark);
  }, [isDark]);

  const toggleTheme = useCallback(() => {
    const newMode: ThemeMode = isDark ? 'light' : 'dark';
    setMode(newMode);
    const settings = loadSettings();
    saveSettings({ ...settings, theme: newMode });
  }, [isDark]);

  const setTheme = useCallback((newMode: ThemeMode) => {
    setMode(newMode);
    const settings = loadSettings();
    saveSettings({ ...settings, theme: newMode });
  }, []);

  // Use custom NVIDIA-branded themes
  const theme: Theme = isDark ? customDarkTheme : customLightTheme;

  return {
    theme,
    mode,
    isDark,
    toggleTheme,
    setTheme,
  };
}
