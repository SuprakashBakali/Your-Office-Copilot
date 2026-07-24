import React, { useState } from 'react';
import { makeStyles, tokens, Card, CardHeader, Text, Input, TabList, Tab, Badge } from '@fluentui/react-components';
import { SearchRegular, StarRegular, Star20Filled } from '@fluentui/react-icons';
import { getTemplatesForHost, searchTemplates, PROMPT_TEMPLATES } from '../../utils/prompts';
import { useAppState, useAppDispatch } from '../../store/AppContext';
import { PromptCategory } from '../../types';
import { loadFavoritePrompts, toggleFavorite as toggleFav } from '../../utils/storage';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    height: '100%',
    padding: '12px',
  },
  searchBox: {
    width: '100%',
  },
  tabArea: {
    overflowX: 'auto',
  },
  grid: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    overflowY: 'auto',
    flexGrow: 1,
  },
  card: {
    cursor: 'pointer',
    flexShrink: 0,
    padding: '8px 4px',
    transition: 'all 0.2s ease',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      transform: 'translateY(-1px)',
      boxShadow: `0 2px 8px ${tokens.colorNeutralShadowAmbient}`,
    },
  },
  iconEmoji: {
    fontSize: '20px',
    width: '28px',
    height: '28px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    borderRadius: '6px',
    backgroundColor: tokens.colorNeutralBackground3,
  },
  favoriteBtn: {
    cursor: 'pointer',
    padding: '4px',
    borderRadius: '4px',
    display: 'flex',
    alignItems: 'center',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    padding: '32px',
    color: tokens.colorNeutralForeground3,
  },
});

interface PromptTemplatesProps {
  onSelect?: (prompt: string) => void;
  hostApp?: string;
}

export const PromptTemplates: React.FC<PromptTemplatesProps> = ({ onSelect, hostApp }) => {
  const classes = useStyles();
  const state = useAppState();
  const dispatch = useAppDispatch();
  const [search, setSearch] = useState('');
  const [category, setCategory] = useState<string>('All');
  const [favorites, setFavorites] = useState<Set<string>>(() => new Set(loadFavoritePrompts()));

  const effectiveHost = hostApp || (state.host === 'Unknown' ? 'Excel' : state.host);

  const templates = search
    ? searchTemplates(search, effectiveHost as any)
    : getTemplatesForHost(effectiveHost as any);

  const filteredTemplates = category === 'All'
    ? templates
    : category === 'Favorites'
      ? templates.filter((t: any) => favorites.has(t.id))
      : templates.filter((t: any) => t.category === category.toLowerCase());

  const handleToggleFavorite = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const updated = toggleFav(id);
    setFavorites(new Set(updated));
  };

  const handleSelect = (prompt: string) => {
    if (onSelect) {
      onSelect(prompt);
    } else {
      // Switch to chat view and populate the input
      dispatch({ type: 'SET_PENDING_PROMPT', payload: prompt });
      dispatch({ type: 'SET_VIEW', payload: 'chat' });
    }
  };

  const categories = ['All', 'Favorites', 'Analysis', 'Formula', 'Chart', 'Finance', 'Cleaning', 'Writing', 'Code', 'Accounting', 'General'];

  return (
    <div className={classes.container}>
      <Input
        className={classes.searchBox}
        placeholder="Search templates..."
        contentBefore={<SearchRegular />}
        value={search}
        onChange={(_, d) => setSearch(d.value)}
      />

      <div className={classes.tabArea}>
        <TabList
          selectedValue={category}
          onTabSelect={(_, d) => setCategory(d.value as string)}
          size="small"
        >
          {categories.map(cat => (
            <Tab key={cat} value={cat}>
              {cat}
              {cat === 'Favorites' && favorites.size > 0 && (
                <Badge size="small" appearance="filled" color="brand" style={{ marginLeft: '4px' }}>{favorites.size}</Badge>
              )}
            </Tab>
          ))}
        </TabList>
      </div>

      <div className={classes.grid}>
        {filteredTemplates.length === 0 ? (
          <div className={classes.emptyState}>
            <Text size={300}>No templates found</Text>
            <Text size={200}>Try a different search or category</Text>
          </div>
        ) : (
          filteredTemplates.map((template: any) => (
            <Card
              key={template.id}
              className={classes.card}
              onClick={() => handleSelect(template.prompt)}
              size="small"
            >
              <CardHeader
                image={<div className={classes.iconEmoji}>{template.icon}</div>}
                header={<Text weight="semibold" size={300}>{template.title}</Text>}
                description={<Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>{template.description}</Text>}
                action={
                  <div className={classes.favoriteBtn} onClick={(e) => handleToggleFavorite(template.id, e)}>
                    {favorites.has(template.id)
                      ? <Star20Filled style={{ color: '#FFB900' }} />
                      : <StarRegular />
                    }
                  </div>
                }
              />
            </Card>
          ))
        )}
      </div>
    </div>
  );
};
