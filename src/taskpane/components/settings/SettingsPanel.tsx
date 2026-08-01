import React, { useState } from 'react';
import {
  makeStyles, tokens, Accordion, AccordionItem, AccordionHeader, AccordionPanel,
  Dropdown, Option, Switch, Button, Text, Slider, Dialog, DialogSurface,
  DialogTitle, DialogBody, DialogActions, DialogTrigger, Divider,
} from '@fluentui/react-components';
import { DeleteRegular, ArrowExportRegular, ArrowResetRegular, InfoRegular } from '@fluentui/react-icons';
import { useSettings } from '../../hooks/useSettings';
import { useTheme } from '../../hooks/useTheme';
import { CustomModelManager } from './CustomModelManager';
import { useAppDispatch } from '../../store/AppContext';
import { clearAllData } from '../../utils/storage';

const useStyles = makeStyles({
  container: {
    padding: '16px',
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    overflowY: 'auto',
    height: '100%',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
  field: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: '12px',
  },
  fieldLabel: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    flexShrink: 0,
  },
  actions: {
    display: 'flex',
    gap: '8px',
    marginTop: '16px',
    flexWrap: 'wrap',
  },
  dangerZone: {
    padding: '12px',
    borderRadius: '8px',
    border: `1px solid ${tokens.colorPaletteRedBorder2}`,
    backgroundColor: tokens.colorPaletteRedBackground1,
  },
  versionInfo: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '12px',
    borderRadius: '8px',
    backgroundColor: tokens.colorNeutralBackground3,
  },
});

export const SettingsPanel: React.FC = () => {
  const classes = useStyles();
  const { settings, updateSettings, resetSettings } = useSettings();
  const { isDark, toggleTheme } = useTheme();
  const dispatch = useAppDispatch();
  const [showClearDialog, setShowClearDialog] = useState(false);

  const handleClearAll = () => {
    clearAllData();
    resetSettings();
    setShowClearDialog(false);
    dispatch({ type: 'SET_NOTIFICATION', payload: { type: 'success', message: 'All data cleared successfully' } });
  };

  const handleExportSettings = () => {
    const data = JSON.stringify(settings, null, 2);
    const blob = new Blob([data], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'ai-copilot-settings.json';
    a.click();
    URL.revokeObjectURL(url);
    dispatch({ type: 'SET_NOTIFICATION', payload: { type: 'success', message: 'Settings exported' } });
  };

  return (
    <div className={classes.container}>
      <Text size={500} weight="semibold">Settings</Text>

      <Accordion multiple defaultOpenItems={['ai', 'keys', 'appearance', 'privacy', 'chat']}>

        {/* My Models */}
        <AccordionItem value="ai">
          <AccordionHeader>🤖 My Models</AccordionHeader>
          <AccordionPanel className={classes.section}>
            <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
              Add models with their own provider and API key. Each model can use a different provider.
            </Text>
            <CustomModelManager />
          </AccordionPanel>
        </AccordionItem>

        {/* Appearance */}
        <AccordionItem value="appearance">
          <AccordionHeader>🎨 Appearance</AccordionHeader>
          <AccordionPanel className={classes.section}>
            <div className={classes.field}>
              <Text>Dark Mode</Text>
              <Switch checked={isDark} onChange={toggleTheme} />
            </div>
            <div className={classes.field}>
              <Text>Font Size</Text>
              <Dropdown
                value={settings.fontSize}
                onOptionSelect={(_, data) => updateSettings({ fontSize: data.optionValue as any })}
                size="small"
              >
                <Option value="small">Small</Option>
                <Option value="medium">Medium</Option>
                <Option value="large">Large</Option>
              </Dropdown>
            </div>
            <div className={classes.field}>
              <Text>Compact Mode</Text>
              <Switch
                checked={settings.compactMode}
                onChange={(_, data) => updateSettings({ compactMode: data.checked })}
              />
            </div>
          </AccordionPanel>
        </AccordionItem>

        {/* Chat Settings */}
        <AccordionItem value="chat">
          <AccordionHeader>💬 Chat</AccordionHeader>
          <AccordionPanel className={classes.section}>
            <div className={classes.field}>
              <div className={classes.fieldLabel}>
                <Text>Stream Responses</Text>
                <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Show text as it generates</Text>
              </div>
              <Switch
                checked={settings.streamResponses}
                onChange={(_, data) => updateSettings({ streamResponses: data.checked })}
              />
            </div>
            <div className={classes.field}>
              <div className={classes.fieldLabel}>
                <Text>Save Conversations</Text>
                <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Persist chat history</Text>
              </div>
              <Switch
                checked={settings.saveConversations}
                onChange={(_, data) => updateSettings({ saveConversations: data.checked })}
              />
            </div>
            <div className={classes.field}>
              <Text>Max History</Text>
              <Slider
                min={10}
                max={200}
                step={10}
                value={settings.maxConversationHistory}
                onChange={(_, data) => updateSettings({ maxConversationHistory: data.value })}
              />
            </div>
          </AccordionPanel>
        </AccordionItem>

        {/* Privacy */}
        <AccordionItem value="privacy">
          <AccordionHeader>🔒 Privacy & Context</AccordionHeader>
          <AccordionPanel className={classes.section}>
            <div className={classes.field}>
              <div className={classes.fieldLabel}>
                <Text>Include Context by Default</Text>
                <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Auto-include Excel data in chat</Text>
              </div>
              <Switch
                checked={settings.includeContextByDefault}
                onChange={(_, data) => updateSettings({ includeContextByDefault: data.checked })}
              />
            </div>
            <div className={classes.field}>
              <div className={classes.fieldLabel}>
                <Text>Max Context Cells</Text>
                <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>{settings.maxContextCells} cells</Text>
              </div>
              <Slider
                min={50}
                max={5000}
                step={50}
                value={settings.maxContextCells}
                onChange={(_, data) => updateSettings({ maxContextCells: data.value })}
              />
            </div>
            <div className={classes.field}>
              <div className={classes.fieldLabel}>
                <Text>Max Context Characters</Text>
                <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>{settings.maxContextCharacters.toLocaleString()} chars</Text>
              </div>
              <Slider
                min={1000}
                max={50000}
                step={1000}
                value={settings.maxContextCharacters}
                onChange={(_, data) => updateSettings({ maxContextCharacters: data.value })}
              />
            </div>
          </AccordionPanel>
        </AccordionItem>
      </Accordion>

      <Divider />

      {/* Actions */}
      <div className={classes.actions}>
        <Button icon={<ArrowExportRegular />} onClick={handleExportSettings}>Export Settings</Button>
        <Button icon={<ArrowResetRegular />} onClick={resetSettings} appearance="subtle">Reset to Defaults</Button>
      </div>

      {/* Danger Zone */}
      <div className={classes.dangerZone}>
        <Text weight="semibold" style={{ color: tokens.colorPaletteRedForeground1, display: 'flex', alignItems: 'center', gap: '6px' }}>
          ⚠️ Danger Zone
        </Text>
        <Text size={200} style={{ color: tokens.colorPaletteRedForeground1, marginBottom: '8px', display: 'block' }}>
          This will delete all conversations, settings, and API keys.
        </Text>
        <Dialog open={showClearDialog} onOpenChange={(_, data) => setShowClearDialog(data.open)}>
          <DialogTrigger>
            <Button icon={<DeleteRegular />} appearance="primary" style={{ backgroundColor: tokens.colorPaletteRedBackground3, borderColor: tokens.colorPaletteRedBorder2 }}>
              Clear All Data
            </Button>
          </DialogTrigger>
          <DialogSurface>
            <DialogTitle>Clear All Data?</DialogTitle>
            <DialogBody>
              This will permanently delete all conversations, settings, API keys, and favorites. This action cannot be undone.
            </DialogBody>
            <DialogActions>
              <DialogTrigger><Button appearance="secondary">Cancel</Button></DialogTrigger>
              <Button appearance="primary" onClick={handleClearAll} style={{ backgroundColor: tokens.colorPaletteRedBackground3 }}>
                Delete Everything
              </Button>
            </DialogActions>
          </DialogSurface>
        </Dialog>
      </div>

      {/* Version Info */}
      <div className={classes.versionInfo}>
        <InfoRegular />
        <div>
          <Text size={200} weight="semibold">Your Co-Pilot v1.0.0</Text>
          <br />
          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Created by Suprakash Bakali</Text>
        </div>
      </div>
    </div>
  );
};
