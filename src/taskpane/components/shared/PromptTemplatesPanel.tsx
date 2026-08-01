import React, { useMemo, useState } from 'react';
import {
  makeStyles, tokens, Popover, PopoverTrigger, PopoverSurface,
  MenuButton,
  Input, Tooltip,
} from '@fluentui/react-components';
import { SparkleRegular, SearchRegular } from '@fluentui/react-icons';
import { useAppState } from '../../store/AppContext';
import {
  getTemplatesForHost, groupTemplatesByCategory,
  type PromptTemplate,
} from '../../services/promptTemplates';

const useStyles = makeStyles({
  trigger: {
    minWidth: 'auto',
  },
  surface: {
    width: '320px',
    maxHeight: '420px',
    overflowY: 'auto',
    padding: '8px',
  },
  search: {
    marginBottom: '8px',
  },
  categoryHeader: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
    margin: '8px 4px 4px',
    fontWeight: 600,
  },
  item: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    padding: '6px 8px',
    borderRadius: '6px',
    cursor: 'pointer',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground2Hover,
    },
  },
  itemTitle: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    fontSize: tokens.fontSizeBase200,
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
  },
  itemDesc: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    paddingLeft: '22px',
  },
  empty: {
    padding: '16px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
});

interface PromptTemplatesPanelProps {
  onPick: (prompt: string) => void;
}

/**
 * Quick-Actions / Prompt Templates panel — opens from a sparkle icon in the
 * nav bar. Curated, host-aware prompt library inspired by office-agents'
 * Skills system and OfficeCLI's specialized skill triggers (pitch-deck,
 * financial-model, dashboard, academic-paper).
 */
export const PromptTemplatesPanel: React.FC<PromptTemplatesPanelProps> = ({ onPick }) => {
  const classes = useStyles();
  const { host } = useAppState();
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');

  const filtered = useMemo(() => {
    const hostTemplates = getTemplatesForHost(host);
    if (!query.trim()) return hostTemplates;
    const q = query.toLowerCase();
    return hostTemplates.filter(t =>
      t.title.toLowerCase().includes(q) ||
      t.description.toLowerCase().includes(q) ||
      t.prompt.toLowerCase().includes(q),
    );
  }, [host, query]);

  const grouped = useMemo(() => groupTemplatesByCategory(filtered), [filtered]);

  const handlePick = (template: PromptTemplate) => {
    onPick(template.prompt);
    setOpen(false);
    setQuery('');
  };

  return (
    <Popover
      open={open}
      onOpenChange={(_, data) => setOpen(data.open)}
      positioning="below"
    >
      <PopoverTrigger>
        <Tooltip content="Prompt Templates" relationship="label">
          <MenuButton
            appearance="subtle"
            size="small"
            icon={<SparkleRegular />}
            className={classes.trigger}
          />
        </Tooltip>
      </PopoverTrigger>
      <PopoverSurface className={classes.surface}>
        <Input
          className={classes.search}
          placeholder="Search templates..."
          value={query}
          onChange={(_, d) => setQuery(d.value)}
          contentBefore={<SearchRegular />}
          size="small"
        />
        {Object.keys(grouped).length === 0 ? (
          <div className={classes.empty}>No templates match &ldquo;{query}&rdquo;</div>
        ) : (
          Object.entries(grouped).map(([category, items]) => (
            <div key={category}>
              <div className={classes.categoryHeader}>{category}</div>
              {items.map(t => (
                <div
                  key={t.id}
                  className={classes.item}
                  onClick={() => handlePick(t)}
                  role="button"
                  tabIndex={0}
                  onKeyDown={e => { if (e.key === 'Enter') handlePick(t); }}
                >
                  <div className={classes.itemTitle}>
                    <span>{t.icon}</span>
                    <span>{t.title}</span>
                  </div>
                  <div className={classes.itemDesc}>{t.description}</div>
                </div>
              ))}
            </div>
          ))
        )}
      </PopoverSurface>
    </Popover>
  );
};
