import React, { useState, useCallback } from 'react';
import { makeStyles, tokens, Text, Button, Badge, Spinner, Tooltip, Textarea } from '@fluentui/react-components';
import { CodeRegular, DocumentRegular, PlugConnectedRegular, DatabaseRegular, DismissRegular } from '@fluentui/react-icons';
import { MarkdownRenderer } from '../shared/MarkdownRenderer';
import { CopyButton } from '../shared/CopyButton';
import { useExcelTools } from '../../hooks/useExcelTools';

const CODE_SYSTEM_PROMPT = `You are an expert code generation assistant for Microsoft Excel automation. You generate clean, well-commented, production-ready code. Always include error handling, comments explaining each step, and instructions on how to use the generated code.`;

const useStyles = makeStyles({
  container: { display: 'flex', flexDirection: 'column', gap: '12px', padding: '16px', height: '100%', overflowY: 'auto' },
  section: {
    display: 'flex', flexDirection: 'column', gap: '8px', padding: '14px', borderRadius: '10px',
    backgroundColor: tokens.colorNeutralBackground1, border: `1px solid ${tokens.colorNeutralStroke2}`,
    '&:hover': { border: `1px solid ${tokens.colorBrandStroke1}` },
  },
  grid: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' },
  card: {
    cursor: 'pointer', padding: '12px', borderRadius: '10px',
    backgroundColor: tokens.colorNeutralBackground1, border: `1px solid ${tokens.colorNeutralStroke2}`,
    transition: 'all 0.2s ease', display: 'flex', flexDirection: 'column', gap: '4px',
    '&:hover': { backgroundColor: tokens.colorNeutralBackground1Hover, border: `1px solid ${tokens.colorBrandStroke1}`, transform: 'translateY(-1px)' },
  },
  resultArea: { padding: '14px', backgroundColor: tokens.colorNeutralBackground3, borderRadius: '10px', border: `1px solid ${tokens.colorNeutralStroke2}` },
  resultHeader: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' },
});

const QUICK_TOOLS = [
  { id: 'vba', title: 'VBA Macro', desc: 'Automate with VBA', icon: <CodeRegular />, lang: 'VBA',
    promptPrefix: 'Generate a VBA macro for Excel that' },
  { id: 'office-script', title: 'Office Script', desc: 'TypeScript automation', icon: <DocumentRegular />, lang: 'TypeScript (Office Script)',
    promptPrefix: 'Generate an Office Script (TypeScript) for Excel that' },
  { id: 'power-query', title: 'Power Query (M)', desc: 'Data transformation', icon: <PlugConnectedRegular />, lang: 'Power Query M',
    promptPrefix: 'Generate Power Query M code that' },
  { id: 'sql', title: 'SQL Query', desc: 'Query generation', icon: <DatabaseRegular />, lang: 'SQL',
    promptPrefix: 'Based on this Excel data structure, generate SQL that' },
  { id: 'python', title: 'Python (pandas)', desc: 'Data analysis code', icon: <CodeRegular />, lang: 'Python',
    promptPrefix: 'Generate Python code using pandas that' },
  { id: 'power-automate', title: 'Power Automate', desc: 'Flow suggestions', icon: <PlugConnectedRegular />, lang: 'Power Automate flow',
    promptPrefix: 'Suggest a Power Automate flow design that' },
];

export const CodeGenTools: React.FC = () => {
  const classes = useStyles();
  const [description, setDescription] = useState('');
  const [selectedLang, setSelectedLang] = useState<string | null>(null);
  const { result, isLoading, error, runTool, clearResult, cancelStream } = useExcelTools();

  const generateCode = useCallback(async (toolId: string) => {
    const tool = QUICK_TOOLS.find(t => t.id === toolId);
    if (!tool) return;
    setSelectedLang(tool.lang);

    const prompt = description.trim()
      ? `${tool.promptPrefix}: ${description}.\n\nInclude: proper error handling, comments explaining each step, and instructions on how to use the code (where to paste it, how to run it).`
      : `${tool.promptPrefix} works with the currently selected Excel data. Analyze the data structure and generate useful ${tool.lang} code for common operations. Include: data loading, transformation, analysis, and output. Add comprehensive comments.`;

    await runTool(
      `${CODE_SYSTEM_PROMPT}\n\nGenerate ${tool.lang} code specifically.`,
      prompt,
      { includeSelection: true, includeWorkbook: true, includeFormulas: true }
    );
  }, [description, runTool]);

  return (
    <div className={classes.container}>
      <Text size={400} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <CodeRegular /> Code Generation
      </Text>

      <div className={classes.section}>
        <Text weight="semibold" size={300}>Describe what you want to automate (optional)</Text>
        <Textarea
          placeholder="e.g., 'Format all dates in column A to YYYY-MM-DD' or 'Create a summary of sales by region'..."
          value={description}
          onChange={(_, d) => setDescription(d.value)}
          resize="vertical"
          style={{ minHeight: '50px' }}
        />
      </div>

      <div className={classes.grid}>
        {QUICK_TOOLS.map(tool => (
          <div key={tool.id} className={classes.card} onClick={() => !isLoading && generateCode(tool.id)}
            style={{ opacity: isLoading ? 0.6 : 1, cursor: isLoading ? 'wait' : 'pointer' }}>
            <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px', color: tokens.colorBrandForeground1 }}>
              {tool.icon} {tool.title}
            </Text>
            <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>{tool.desc}</Text>
          </div>
        ))}
      </div>

      {isLoading && !result && (
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '16px', justifyContent: 'center' }}>
          <Spinner size="small" /><Text size={200}>Generating {selectedLang} code...</Text>
          <Button size="small" appearance="subtle" onClick={cancelStream}>Cancel</Button>
        </div>
      )}

      {(result || error) && (
        <div className={classes.resultArea}>
          <div className={classes.resultHeader}>
            <Badge appearance="outline" color={error ? 'danger' : 'success'} size="small">{error ? '⚠️ Error' : `⚙️ ${selectedLang || 'Code'}`}</Badge>
            <div style={{ display: 'flex', gap: '2px' }}>
              <CopyButton text={result} />
              <Tooltip content="Clear" relationship="label"><Button appearance="transparent" size="small" icon={<DismissRegular />} onClick={clearResult} /></Tooltip>
            </div>
          </div>
          {error ? <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{error}</Text> : <MarkdownRenderer content={result} />}
        </div>
      )}
    </div>
  );
};
