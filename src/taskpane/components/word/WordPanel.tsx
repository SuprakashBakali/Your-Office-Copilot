import React, { useState, useCallback } from 'react';
import { makeStyles, tokens, Card, CardHeader, Text, Button, Badge, Spinner, Tooltip } from '@fluentui/react-components';
import {
  DocumentTextRegular, EditRegular, CheckboxCheckedRegular, MailRegular,
  ExpandUpRightRegular, DismissRegular, TextCaseTitleRegular,
} from '@fluentui/react-icons';
import { MarkdownRenderer } from '../shared/MarkdownRenderer';
import { CopyButton } from '../shared/CopyButton';
import { useAI } from '../../hooks/useAI';
import { useSettings } from '../../hooks/useSettings';
import { ChatMessage } from '../../types';

const WORD_SYSTEM_PROMPT = `You are a professional writing assistant integrated into Microsoft Word. Help users rewrite, summarise, proofread, and draft content. Always maintain the original meaning unless told otherwise. Format your responses using Markdown.`;

const useStyles = makeStyles({
  container: {
    padding: '16px',
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    height: '100%',
    overflowY: 'auto',
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: '8px',
  },
  card: {
    cursor: 'pointer',
    padding: '12px',
    borderRadius: '10px',
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    transition: 'all 0.2s ease',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      border: `1px solid ${tokens.colorBrandStroke1}`,
      transform: 'translateY(-1px)',
      boxShadow: `0 2px 8px ${tokens.colorNeutralShadowAmbient}`,
    },
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
  actions: {
    display: 'flex',
    gap: '8px',
    marginTop: '12px',
  },
});

const WORD_TOOLS = [
  {
    id: 'rewrite',
    title: 'Rewrite Selection',
    desc: 'Professional rewriting',
    icon: <EditRegular />,
    prompt: 'Rewrite the following text to be more professional, clear, and concise. Maintain the original meaning but improve the language, structure, and flow. Provide the rewritten version.',
  },
  {
    id: 'summarize',
    title: 'Summarize Document',
    desc: 'Concise summary',
    icon: <DocumentTextRegular />,
    prompt: 'Summarize this text concisely. Provide: a one-paragraph executive summary, followed by key points as bullet points. Keep it under 200 words.',
  },
  {
    id: 'grammar',
    title: 'Grammar Check',
    desc: 'Fix grammar & style',
    icon: <CheckboxCheckedRegular />,
    prompt: 'Review this text for grammar, spelling, punctuation, and style errors. For each issue: quote the problematic text, explain the error, and provide the correction. Then provide the fully corrected version.',
  },
  {
    id: 'expand',
    title: 'Expand Content',
    desc: 'Add more detail',
    icon: <ExpandUpRightRegular />,
    prompt: 'Expand this text by adding more detail, examples, and depth while maintaining the same tone and style. Aim for roughly 2-3x the original length.',
  },
  {
    id: 'email',
    title: 'Draft Email',
    desc: 'Professional email',
    icon: <MailRegular />,
    prompt: 'Draft a professional email based on this context. Include: subject line, greeting, clear body with purpose, action items, and a professional closing. Keep it concise.',
  },
  {
    id: 'formal',
    title: 'Make Formal',
    desc: 'Business tone',
    icon: <TextCaseTitleRegular />,
    prompt: 'Rewrite this text in a formal, professional business tone suitable for official correspondence, reports, or executive communication. Maintain accuracy.',
  },
];

export const WordPanel: React.FC = () => {
  const classes = useStyles();
  const [result, setResult] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [activeTool, setActiveTool] = useState<string | null>(null);
  const ai = useAI();
  const { settings } = useSettings();

  const handleAction = useCallback(async (toolId: string) => {
    const tool = WORD_TOOLS.find(t => t.id === toolId);
    if (!tool) return;

    setIsLoading(true);
    setError(null);
    setResult('');
    setActiveTool(tool.title);

    try {
      // Try to get Word document context
      let contextStr = '';
      try {
        const { WordService } = await import('../../services/office/WordService');
        contextStr = await WordService.getContextForAI();
      } catch {
        contextStr = '[Note: Could not read Word document — please describe or paste your text in the chat.]';
      }

      const systemPrompt = `${WORD_SYSTEM_PROMPT}\n\nCurrent document context:\n${contextStr}`;
      const messages: ChatMessage[] = [
        { id: 'sys', role: 'system', content: systemPrompt, timestamp: 0 },
        { id: 'usr', role: 'user', content: tool.prompt, timestamp: Date.now() },
      ];

      let response = '';

      if (settings.streamResponses) {
        response = await ai.sendMessageStream(messages, {}, (chunk) => {
          setResult(prev => prev + chunk);
        });
      } else {
        response = await ai.sendMessage(messages);
        setResult(response);
      }

      setIsLoading(false);
    } catch (err) {
      setError((err as Error).message);
      setIsLoading(false);
    }
  }, [ai]);

  const clearResult = useCallback(() => {
    setResult('');
    setError(null);
    setActiveTool(null);
  }, []);

  return (
    <div className={classes.container}>
      <Text size={400} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <DocumentTextRegular /> Word Tools
      </Text>
      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
        Select text in Word, then choose a tool
      </Text>

      <div className={classes.grid}>
        {WORD_TOOLS.map(tool => (
          <div
            key={tool.id}
            className={classes.card}
            onClick={() => !isLoading && handleAction(tool.id)}
            style={{ opacity: isLoading ? 0.6 : 1, cursor: isLoading ? 'wait' : 'pointer' }}
          >
            <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px', color: tokens.colorBrandForeground1 }}>
              {tool.icon} {tool.title}
            </Text>
            <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>{tool.desc}</Text>
          </div>
        ))}
      </div>

      {isLoading && !result && (
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '16px', justifyContent: 'center' }}>
          <Spinner size="small" />
          <Text size={200}>Processing with AI...</Text>
          <Button size="small" appearance="subtle" onClick={ai.cancelStream}>Cancel</Button>
        </div>
      )}

      {(result || error) && (
        <div className={classes.resultArea}>
          <div className={classes.resultHeader}>
            <Badge appearance="outline" color={error ? 'danger' : 'success'} size="small">
              {error ? '⚠️ Error' : `✏️ ${activeTool || 'Result'}`}
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
