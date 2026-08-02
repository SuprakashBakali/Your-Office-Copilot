import React, { useState, useCallback, useRef, useEffect } from 'react';
import { makeStyles, tokens, Text, Button, Badge, Spinner, Tooltip, Textarea } from '@fluentui/react-components';
import {
  SlideTextRegular, NoteRegular, TextBulletListLtrRegular,
  ChatHelpRegular, PersonBoardRegular, DismissRegular,
  ColorRegular, TextFontRegular, PlayCircleRegular, ImageRegular, ShapesRegular, TimelineRegular, DocumentTextRegular, OrganizationRegular, ArrowRoutingRegular, TableRegular
} from '@fluentui/react-icons';
import { MarkdownRenderer } from '../shared/MarkdownRenderer';
import { CopyButton } from '../shared/CopyButton';
import { useAI } from '../../hooks/useAI';
import { useSettings } from '../../hooks/useSettings';
import { ChatMessage } from '../../types';

const PPT_SYSTEM_PROMPT = `You are a professional presentation assistant integrated into Microsoft PowerPoint. Help users create, improve, and polish their slides and speaker notes. Use concise, impactful language suitable for presentations. Format your responses using Markdown.`;

const useStyles = makeStyles({
  container: {
    padding: '16px',
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    height: '100%',
    overflowY: 'auto',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    padding: '14px',
    borderRadius: '10px',
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    '&:hover': {
      border: `1px solid ${tokens.colorBrandStroke1}`,
    },
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
});

const PPT_TOOLS = [
  {
    id: 'improve',
    title: 'Improve Slide',
    desc: 'Make text impactful',
    icon: <SlideTextRegular />,
    prompt: 'Improve this slide text to be more impactful and presentation-ready. Use: concise bullet points, active voice, powerful verbs, and clear messaging. Keep each bullet under 10 words.',
  },
  {
    id: 'notes',
    title: 'Speaker Notes',
    desc: 'Generate talking points',
    icon: <NoteRegular />,
    prompt: 'Generate detailed speaker notes for this slide content. Include: what to say, key talking points, transitions to the next slide, and audience engagement tips. Write in a natural speaking style.',
  },
  {
    id: 'outline',
    title: 'Slide Outline',
    desc: 'Create structure',
    icon: <TextBulletListLtrRegular />,
    prompt: 'Create a professional presentation outline from this content. Include: title slide, agenda, 5-8 content slides with bullet points, key visuals to include, and a closing slide. Format as slide-by-slide with titles and bullets.',
  },
  {
    id: 'qa',
    title: 'Q&A Prep',
    desc: 'Anticipate questions',
    icon: <ChatHelpRegular />,
    prompt: 'Based on this presentation content, prepare for Q&A. Generate: 10 likely questions the audience might ask, with recommended answers for each. Include tough/challenging questions too.',
  },
  {
    id: 'board',
    title: 'Board Ready',
    desc: 'Executive audience',
    icon: <PersonBoardRegular />,
    prompt: 'Adapt this slide content for a board/executive meeting. Focus on: strategic impact, KPIs, financial implications, risks, and recommendations. Use concise, data-driven language. Remove operational details.',
  },
  { id: 'design-ideas', title: 'Design Advisor', desc: 'Visual improvements', icon: <ColorRegular />, prompt: 'Suggest 3 specific design improvements for this slide to make it more visually engaging. Focus on contrast, alignment, whitespace, and visual hierarchy.' },
  { id: 'layout', title: 'Layout Optimizer', desc: 'Structure slide', icon: <SlideTextRegular />, prompt: 'Suggest the optimal slide layout for this content (e.g., Two Content, Title Only, Picture with Caption). Explain how to arrange the text and visuals for maximum impact.' },
  { id: 'color', title: 'Color Palette', desc: 'Brand colors', icon: <ColorRegular />, prompt: 'Recommend a professional 5-color palette (with HEX codes) for this presentation. Include a primary, secondary, accent, dark text, and light background color.' },
  { id: 'fonts', title: 'Font Advisor', desc: 'Typography', icon: <TextFontRegular />, prompt: 'Suggest a professional font pairing (heading and body) for this presentation. Explain why they work well together and suggest optimal font sizes.' },
  { id: 'animation', title: 'Animation Guide', desc: 'Slide transitions', icon: <PlayCircleRegular />, prompt: 'Suggest professional animations and slide transitions for this content. Avoid distracting effects. Focus on guiding the viewer\'s eye (e.g., Fade, Push, Wipe).' },
  { id: 'image', title: 'Image Ideas', desc: 'Visual metaphors', icon: <ImageRegular />, prompt: 'Suggest 3 high-quality, professional image concepts or visual metaphors that would complement this text perfectly. Describe what to search for.' },
  { id: 'icons', title: 'Icon Matcher', desc: 'Vector graphics', icon: <ShapesRegular />, prompt: 'Suggest 5 specific, simple icons (e.g., "Target", "Lightbulb", "Gears") that map to the key concepts in this text.' },
  { id: 'timeline', title: 'Timeline Builder', desc: 'Chronological', icon: <TimelineRegular />, prompt: 'Convert this text into a logical timeline structure for a slide. Organize it by dates/phases and summarize the milestones.' },
  { id: 'swot-ppt', title: 'SWOT Slide', desc: 'Strengths & Weaknesses', icon: <DocumentTextRegular />, prompt: 'Format this content into a 2x2 SWOT analysis matrix (Strengths, Weaknesses, Opportunities, Threats) suitable for a slide layout.' },
  { id: 'pestle', title: 'PESTLE Slide', desc: 'Macro environment', icon: <DocumentTextRegular />, prompt: 'Format this content into a PESTLE analysis (Political, Economic, Social, Technological, Legal, Environmental) with short, punchy bullet points.' },
  { id: 'canvas', title: 'Business Canvas', desc: 'Business model', icon: <DocumentTextRegular />, prompt: 'Map this content to the 9 blocks of the Business Model Canvas. Keep the text extremely concise for a single slide.' },
  { id: 'roadmap', title: 'Roadmap Creator', desc: 'Future plan', icon: <TimelineRegular />, prompt: 'Translate this strategic text into a structured roadmap (Now, Next, Later) with clear deliverables for each phase.' },
  { id: 'orgchart', title: 'Org Chart', desc: 'Hierarchy structure', icon: <OrganizationRegular />, prompt: 'Extract the reporting structure and hierarchy from this text to help me build an Org Chart slide. List roles and who they report to.' },
  { id: 'process', title: 'Process Flow', desc: 'Step-by-step', icon: <ArrowRoutingRegular />, prompt: 'Convert this descriptive text into a clear, sequential step-by-step process flow (Step 1, Step 2...) suitable for a flowchart diagram.' },
  { id: 'compare', title: 'Comparison Table', desc: 'Vs analysis', icon: <TableRegular />, prompt: 'Convert this text into a clear comparison table (e.g., Option A vs Option B) highlighting the pros, cons, and key differences across standard criteria.' }
];

export const PowerPointPanel: React.FC = () => {
  const classes = useStyles();
  const [result, setResult] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [activeTool, setActiveTool] = useState<string | null>(null);
  const [customTopic, setCustomTopic] = useState('');
  const ai = useAI();
  const { settings } = useSettings();

  // H2: prevent setState on unmounted component when user switches tabs mid-stream.
  const mounted = useRef(true);
  // M9: synchronous guard against double-click race — set before the async
  // chain begins, so a second click arriving before setIsLoading(true) fires
  // is rejected immediately.
  const isRunningRef = useRef(false);

  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      // H2: abort any in-flight stream when the panel unmounts.
      ai.cancelStream();
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const runTool = useCallback(async (toolId: string) => {
    const tool = PPT_TOOLS.find(t => t.id === toolId);
    if (!tool) return;
    // M9: close the double-click race window before any async work
    if (isRunningRef.current) return;
    isRunningRef.current = true;

    if (mounted.current) {
      setIsLoading(true);
      setError(null);
      setResult('');
      setActiveTool(tool.title);
    }

    try {
      let contextStr = '';
      try {
        const { PowerPointService } = await import('../../services/office/PowerPointService');
        contextStr = await PowerPointService.getContextForAI();
      } catch {
        contextStr = '[Note: Could not read PowerPoint data — please describe your slide content.]';
      }

      const systemPrompt = `${PPT_SYSTEM_PROMPT}\n\nCurrent slide context:\n${contextStr}`;
      const messages: ChatMessage[] = [
        { id: 'sys', role: 'system', content: systemPrompt, timestamp: 0 },
        { id: 'usr', role: 'user', content: tool.prompt, timestamp: Date.now() },
      ];

      let response = '';

      if (settings.streamResponses) {
        response = (await ai.sendMessageStream(messages, {}, (chunk) => {
          // H2: guard against setState on unmounted component
          if (mounted.current) setResult(prev => prev + chunk);
        })).text;
      } else {
        response = await ai.sendMessage(messages);
        if (mounted.current) setResult(response);
      }

      if (mounted.current) setIsLoading(false);
    } catch (err) {
      if (mounted.current) {
        setError((err as Error).message);
        setIsLoading(false);
      }
    } finally {
      // M9: always clear the guard so the next click is not blocked
      isRunningRef.current = false;
    }
  }, [ai, settings]);

  const generateSlideContent = useCallback(async () => {
    if (!customTopic.trim()) return;
    // M9: same double-click guard for the Generate Slide button
    if (isRunningRef.current) return;
    isRunningRef.current = true;

    if (mounted.current) {
      setIsLoading(true);
      setError(null);
      setResult('');
      setActiveTool('Generate Slide');
    }

    try {
      const messages: ChatMessage[] = [
        { id: 'sys', role: 'system', content: PPT_SYSTEM_PROMPT, timestamp: 0 },
        {
          id: 'usr',
          role: 'user',
          content: `Generate professional slide content for the following topic. Include: a compelling title, 3-5 concise bullet points, and suggested speaker notes.\n\nTopic: ${customTopic}`,
          timestamp: Date.now(),
        },
      ];

      let response = '';

      if (settings.streamResponses) {
        response = (await ai.sendMessageStream(messages, {}, (chunk) => {
          if (mounted.current) setResult(prev => prev + chunk);
        })).text;
      } else {
        response = await ai.sendMessage(messages);
        if (mounted.current) setResult(response);
      }

      if (mounted.current) setIsLoading(false);
    } catch (err) {
      if (mounted.current) {
        setError((err as Error).message);
        setIsLoading(false);
      }
    } finally {
      isRunningRef.current = false;
    }
  }, [customTopic, ai, settings]);

  const clearResult = useCallback(() => {
    setResult('');
    setError(null);
    setActiveTool(null);
  }, []);

  return (
    <div className={classes.container}>
      <Text size={400} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <SlideTextRegular /> PowerPoint Tools
      </Text>

      {/* Slide Content Generator */}
      <div className={classes.section}>
        <Text weight="semibold" size={300}>Generate Slide Content</Text>
        <Textarea
          placeholder="Enter a topic, e.g., 'Q3 Revenue Highlights' or 'Benefits of Cloud Migration'..."
          value={customTopic}
          onChange={(_, d) => setCustomTopic(d.value)}
          resize="vertical"
          style={{ minHeight: '50px' }}
        />
        <Button
          appearance="primary"
          onClick={generateSlideContent}
          disabled={!customTopic.trim() || isLoading}
          icon={isLoading && activeTool === 'Generate Slide' ? <Spinner size="tiny" /> : undefined}
        >
          {isLoading && activeTool === 'Generate Slide' ? 'Generating...' : 'Generate'}
        </Button>
      </div>

      {/* Quick Tools */}
      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
        Or use a quick tool on the current slide
      </Text>
      <div className={classes.grid}>
        {PPT_TOOLS.map(tool => (
          <div
            key={tool.id}
            className={classes.card}
            onClick={() => !isLoading && runTool(tool.id)}
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
          <Text size={200}>Generating with AI...</Text>
          <Button size="small" appearance="subtle" onClick={ai.cancelStream}>Cancel</Button>
        </div>
      )}

      {(result || error) && (
        <div className={classes.resultArea}>
          <div className={classes.resultHeader}>
            <Badge appearance="outline" color={error ? 'danger' : 'success'} size="small">
              {error ? '⚠️ Error' : `🎯 ${activeTool || 'Result'}`}
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
