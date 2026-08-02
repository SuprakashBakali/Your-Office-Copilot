import React, { useState, useCallback } from 'react';
import { makeStyles, Dropdown, Option, Button, Text, tokens, Spinner, Badge, Tooltip } from '@fluentui/react-components';
import { DismissRegular } from '@fluentui/react-icons';
import { MarkdownRenderer } from '../shared/MarkdownRenderer';
import { CopyButton } from '../shared/CopyButton';
import { useAI } from '../../hooks/useAI';
import { useSettings } from '../../hooks/useSettings';
import { ChatMessage } from '../../types';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
    padding: '16px',
  },
  controls: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
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

export const WritingTools: React.FC = () => {
  const classes = useStyles();
  const [tone, setTone] = useState('Professional');
  const [length, setLength] = useState('Same');
  const [result, setResult] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const ai = useAI();
  const { settings } = useSettings();

  const handleRewrite = useCallback(async () => {
    setIsLoading(true);
    setError(null);
    setResult('');

    try {
      let contextStr = '';
      try {
        const { WordService } = await import('../../services/office/WordService');
        contextStr = await WordService.getContextForAI();
      } catch {
        contextStr = '[Note: Could not read Word document — please describe or paste your text.]';
      }

      const systemPrompt = `You are a professional writing assistant. Rewrite the given text according to the user's specifications. Format your response using Markdown.\n\nCurrent document context:\n${contextStr}`;
      const userPrompt = `Rewrite the selected text with the following specifications:\n- Tone: ${tone}\n- Length: ${length === 'Shorter' ? 'Make it about 50% shorter while preserving key information' : length === 'Longer' ? 'Expand to roughly 2x the length with more detail and examples' : 'Keep approximately the same length'}\n\nProvide the rewritten version.`;

      const messages: ChatMessage[] = [
        { id: 'sys', role: 'system', content: systemPrompt, timestamp: 0 },
        { id: 'usr', role: 'user', content: userPrompt, timestamp: Date.now() },
      ];

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
  }, [tone, length, ai]);

  const clearResult = useCallback(() => {
    setResult('');
    setError(null);
  }, []);

  return (
    <div className={classes.container}>
      <Text size={500} weight="semibold">Advanced Writing</Text>
      
      <div className={classes.controls}>
        <Text>Tone</Text>
        <Dropdown value={tone} onOptionSelect={(e, d) => setTone(d.optionValue as string)}>
          <Option value="Professional">Professional</Option>
          <Option value="Casual">Casual</Option>
          <Option value="Academic">Academic</Option>
          <Option value="Creative">Creative</Option>
        </Dropdown>

        <Text>Length</Text>
        <Dropdown value={length} onOptionSelect={(e, d) => setLength(d.optionValue as string)}>
          <Option value="Shorter">Shorter</Option>
          <Option value="Same">Same</Option>
          <Option value="Longer">Longer</Option>
        </Dropdown>

        <Button
          appearance="primary"
          onClick={handleRewrite}
          disabled={isLoading}
          icon={isLoading ? <Spinner size="tiny" /> : undefined}
        >
          {isLoading ? 'Rewriting...' : 'Rewrite Selection'}
        </Button>
      </div>

      {(result || error) && (
        <div className={classes.resultArea}>
          <div className={classes.resultHeader}>
            <Badge appearance="outline" color={error ? 'danger' : 'success'} size="small">
              {error ? '⚠️ Error' : '✏️ Rewritten'}
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
