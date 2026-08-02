import React, { useCallback } from 'react';
import { makeStyles, tokens, Text, Button, Badge, Spinner, Tooltip } from '@fluentui/react-components';
import { DataBarVerticalRegular, ChartMultipleRegular, PaintBrushRegular, SparkleRegular, DismissRegular } from '@fluentui/react-icons';
import { MarkdownRenderer } from '../shared/MarkdownRenderer';
import { CopyButton } from '../shared/CopyButton';
import { useExcelTools } from '../../hooks/useExcelTools';

const CHART_SYSTEM_PROMPT = `You are an expert data visualization specialist working inside Microsoft Excel. You recommend the best chart types, explain how to create them, and suggest formatting. When recommending charts, consider the data structure, story, and audience. Always provide step-by-step Excel instructions.`;

const useStyles = makeStyles({
  container: { display: 'flex', flexDirection: 'column', gap: '12px', padding: '16px', height: '100%', overflowY: 'auto' },
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

const TOOLS = [
  { id: 'recommend', title: 'Recommend Chart', desc: 'Best chart type for your data', icon: <SparkleRegular />,
    prompt: 'Based on this data, recommend the top 3 best chart types. For each: explain why it\'s suitable, what story it tells, and provide step-by-step instructions to create it in Excel. Include formatting tips.' },
  { id: 'dashboard', title: 'Dashboard Design', desc: 'Layout suggestions', icon: <ChartMultipleRegular />,
    prompt: 'Design an Excel dashboard layout for this data. Include: recommended KPI cards, chart placements, filter slicers, and conditional formatting. Provide step-by-step creation instructions.' },
  { id: 'conditional', title: 'Conditional Formatting', desc: 'Visual rules for data', icon: <PaintBrushRegular />,
    prompt: 'Suggest conditional formatting rules for this data. Include: color scales, data bars, icon sets, and custom rules. For each rule: explain the purpose and provide step-by-step instructions. Include VBA for complex rules.' },
  { id: 'pivottable', title: 'PivotTable Design', desc: 'Summarize and group data', icon: <DataBarVerticalRegular />,
    prompt: 'Design a PivotTable for this data. Suggest: which fields to use as rows, columns, values, and filters. Recommend calculated fields and show-values-as options. Provide step-by-step creation instructions.' },
  { id: 'sparklines', title: 'Sparklines', desc: 'In-cell mini charts', icon: <DataBarVerticalRegular />,
    prompt: 'Suggest sparkline configurations for this data. Recommend: which rows/columns to add sparklines to, the type (line/column/win-loss), and formatting options. Provide Excel formulas and instructions.' },
  { id: 'pivotchart', title: 'PivotChart', desc: 'Interactive chart from pivot', icon: <ChartMultipleRegular />,
    prompt: 'Design a PivotChart for this data. Suggest the best chart type, field arrangement, and interactive elements (slicers, timeline). Provide step-by-step creation instructions.' },
];

export const ChartTools: React.FC = () => {
  const classes = useStyles();
  const { result, isLoading, error, runTool, clearResult, cancelStream } = useExcelTools();

  const handleTool = useCallback(async (toolId: string) => {
    const tool = TOOLS.find(t => t.id === toolId);
    if (!tool) return;
    await runTool(CHART_SYSTEM_PROMPT, tool.prompt, { includeSelection: true, includeWorkbook: true });
  }, [runTool]);

  return (
    <div className={classes.container}>
      <Text size={400} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <DataBarVerticalRegular /> Charts & Dashboards
      </Text>
      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Select your data range, then choose a tool</Text>

      <div className={classes.grid}>
        {TOOLS.map(tool => (
          <div key={tool.id} className={classes.card} onClick={() => !isLoading && handleTool(tool.id)}
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
          <Spinner size="small" /><Text size={200}>Creating visualization plan...</Text>
          <Button size="small" appearance="subtle" onClick={cancelStream}>Cancel</Button>
        </div>
      )}

      {(result || error) && (
        <div className={classes.resultArea}>
          <div className={classes.resultHeader}>
            <Badge appearance="outline" color={error ? 'danger' : 'success'} size="small">{error ? '⚠️ Error' : '📉 Chart Recommendations'}</Badge>
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
