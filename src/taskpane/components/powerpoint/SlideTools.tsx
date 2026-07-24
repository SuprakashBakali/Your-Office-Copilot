import React, { useState, useCallback } from 'react';
import { makeStyles, Button, Text, tokens, Textarea, Spinner, Badge, Tooltip } from '@fluentui/react-components';
import { DismissRegular, SlideTextRegular } from '@fluentui/react-icons';
import { MarkdownRenderer } from '../shared/MarkdownRenderer';
import { CopyButton } from '../shared/CopyButton';
import { useAI } from '../../hooks/useAI';
import { loadSettings } from '../../utils/storage';
import { ChatMessage } from '../../types';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
    padding: '16px',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  resultArea: {
    padding: '14px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: '10px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  resultHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '8px',
  },
});

export const SlideTools: React.FC = () => {
  const classes = useStyles();
  const [topic, setTopic] = useState('');
  const [result, setResult] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const ai = useAI();

  const generateSlide = useCallback(async () => {
    if (!topic.trim()) return;

    setIsLoading(true);
    setError(null);
    setResult('');

    try {
      const systemPrompt = `You are a professional presentation designer. Generate compelling slide content from the given topic or notes. Use concise bullet points, strong headers, and clear messaging suitable for professional presentations. Format your response using Markdown.`;
      const userPrompt = `Generate professional slide content for the following topic or notes. Include:\n- A compelling slide title\n- 3-5 concise bullet points (under 10 words each)\n- Suggested visual elements or icons\n- Speaker notes for the presenter\n\nTopic/Notes: ${topic}`;

      const messages: ChatMessage[] = [
        { id: 'sys', role: 'system', content: systemPrompt, timestamp: 0 },
        { id: 'usr', role: 'user', content: userPrompt, timestamp: Date.now() },
      ];

      const settings = loadSettings();

      if (settings.streamResponses) {
        await ai.sendMessageStream(messages, {}, (chunk) => {
          setResult(prev => prev + chunk);
        });
      } else {
        const response = await ai.sendMessage(messages);
        setResult(response);
      }

      setIsLoading(false);
    } catch (err) {
      setError((err as Error).message);
      setIsLoading(false);
    }
  }, [topic, ai]);

  const clearResult = useCallback(() => {
    setResult('');
    setError(null);
  }, []);

  return (
    <div className={classes.container}>
      <Text size={500} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <SlideTextRegular /> Slide Generator
      </Text>
      
      <div className={classes.section}>
        <Text>Topic or Notes</Text>
        <Textarea 
          placeholder="What should this slide be about? e.g., 'Q3 revenue growth of 15% driven by new product launches'"
          value={topic}
          onChange={(e, d) => setTopic(d.value)}
          resize="vertical"
          style={{ minHeight: '60px' }}
        />
        <Button
          appearance="primary"
          onClick={generateSlide}
          disabled={!topic.trim() || isLoading}
          icon={isLoading ? <Spinner size="tiny" /> : undefined}
        >
          {isLoading ? 'Generating...' : 'Generate Content'}
        </Button>
      </div>

      {(result || error) && (
        <div className={classes.resultArea}>
          <div className={classes.resultHeader}>
            <Badge appearance="outline" color={error ? 'danger' : 'success'} size="small">
              {error ? '⚠️ Error' : '🎯 Generated Slide'}
            </Badge>
            <div style={{ display: 'flex', gap: '2px' }}>
              <CopyButton text={result} />
              <Tooltip content="Clear" relationship="label">
                <Button appearance="transparent" size="small" icon={<DismissRegular />} onClick={clearResult} />
              </Tooltip>
            </div>
          </div>
          {error ? (
            <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{error}</Text>
          ) : (
            <MarkdownRenderer content={result} />
          )}
        </div>
      )}
    </div>
  );
};
