import React, { useCallback } from 'react';
import { makeStyles, tokens, Text, Button, Badge, Spinner, Tooltip } from '@fluentui/react-components';
import {
  DataAreaRegular, DataTrendingRegular, SearchRegular,
  CheckmarkSquareRegular, DocumentBulletListRegular, ArrowTrendingRegular,
  LinkRegular, DismissRegular,
} from '@fluentui/react-icons';
import { MarkdownRenderer } from '../shared/MarkdownRenderer';
import { CopyButton } from '../shared/CopyButton';
import { useExcelTools } from '../../hooks/useExcelTools';

const DATA_ANALYSIS_PROMPT = `You are an expert data analyst working inside Microsoft Excel. Analyze the provided data and give actionable, structured insights. Always use Markdown tables, bullet points, and clear headings. Be specific with numbers and statistics. When appropriate, suggest Excel formulas the user can apply.`;

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    padding: '16px',
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
  cardIcon: {
    color: tokens.colorBrandForeground1,
    fontSize: '18px',
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
  loadingOverlay: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '16px',
    justifyContent: 'center',
  },
});

const TOOLS = [
  {
    id: 'stats',
    title: 'Statistical Summary',
    description: 'Count, sum, avg, min, max, median, std dev',
    icon: <DataAreaRegular />,
    prompt: 'Provide a comprehensive statistical summary of this data: count, sum, average, min, max, median, mode, standard deviation, variance, and quartiles. Format as a table. Highlight any notable patterns.',
  },
  {
    id: 'trends',
    title: 'Trend Analysis',
    description: 'Direction, growth rate, patterns',
    icon: <DataTrendingRegular />,
    prompt: 'Analyze trends in this data. Identify: direction (increasing/decreasing/stable), rate of change, seasonal patterns, cyclical behavior, and anomalies. Suggest Excel formulas for trend lines (FORECAST, TREND, GROWTH). Provide forecasting suggestions.',
  },
  {
    id: 'outliers',
    title: 'Outlier Detection',
    description: 'IQR and Z-score analysis',
    icon: <SearchRegular />,
    prompt: 'Detect outliers in this data using IQR method (1.5×IQR rule) and Z-score analysis (|z| > 2). For each outlier: state its value, location, and why it\'s flagged. Suggest how to handle them (remove, cap, investigate).',
  },
  {
    id: 'dupes',
    title: 'Duplicate Detection',
    description: 'Find repeated values and rows',
    icon: <CheckmarkSquareRegular />,
    prompt: 'Scan this data for duplicates. Report: which values/rows appear more than once, their exact locations, frequency counts. Suggest formulas to highlight duplicates (COUNTIF) and remove them.',
  },
  {
    id: 'missing',
    title: 'Missing Values',
    description: 'Find gaps and suggest fills',
    icon: <DocumentBulletListRegular />,
    prompt: 'Analyze this data for missing values, blanks, and gaps. Report: which cells are empty, the percentage of missing data per column, patterns in missingness. Suggest strategies to handle them (fill with mean/median, forward fill, delete rows, interpolate) with Excel formulas.',
  },
  {
    id: 'correlation',
    title: 'Correlation Analysis',
    description: 'Relationships between columns',
    icon: <LinkRegular />,
    prompt: 'Analyze correlations between the columns/variables in this data. For each pair: state the correlation direction (positive/negative), estimated strength (strong/moderate/weak), and practical significance. Suggest Excel formulas (CORREL) and visualization approaches.',
  },
  {
    id: 'forecast',
    title: 'Forecast Suggestions',
    description: 'Predict future values',
    icon: <ArrowTrendingRegular />,
    prompt: 'Based on this data, provide forecasting suggestions. Recommend the best forecasting method (linear trend, moving average, exponential smoothing). Provide Excel formulas (FORECAST.LINEAR, FORECAST.ETS) and explain the expected accuracy.',
  },
  {
    id: 'kpi',
    title: 'KPI Analysis',
    description: 'Key performance indicators',
    icon: <DataAreaRegular />,
    prompt: 'Identify key performance indicators (KPIs) from this data. For each KPI: state the current value, trend, benchmark comparison (if inferable), and RAG status (Red/Amber/Green). Suggest a dashboard layout with formulas.',
  },
];

export const DataTools: React.FC = () => {
  const classes = useStyles();
  const { result, isLoading, error, runTool, clearResult, cancelStream } = useExcelTools();

  const handleTool = useCallback(async (toolId: string) => {
    const tool = TOOLS.find(t => t.id === toolId);
    if (!tool) return;
    await runTool(DATA_ANALYSIS_PROMPT, tool.prompt, {
      includeSelection: true,
      includeFormulas: true,
      includeWorkbook: true,
    });
  }, [runTool]);

  return (
    <div className={classes.container}>
      <Text size={400} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <DataAreaRegular /> Data Analysis
      </Text>
      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
        Select cells in Excel, then click a tool to analyze
      </Text>

      <div className={classes.grid}>
        {TOOLS.map(tool => (
          <div
            key={tool.id}
            className={classes.card}
            onClick={() => !isLoading && handleTool(tool.id)}
            style={{ opacity: isLoading ? 0.6 : 1, cursor: isLoading ? 'wait' : 'pointer' }}
          >
            <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <span className={classes.cardIcon}>{tool.icon}</span>
              {tool.title}
            </Text>
            <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>{tool.description}</Text>
          </div>
        ))}
      </div>

      {isLoading && !result && (
        <div className={classes.loadingOverlay}>
          <Spinner size="small" />
          <Text size={200}>Analyzing your data...</Text>
          <Button size="small" appearance="subtle" onClick={cancelStream}>Cancel</Button>
        </div>
      )}

      {(result || error) && (
        <div className={classes.resultArea}>
          <div className={classes.resultHeader}>
            <Badge appearance="outline" color={error ? 'danger' : 'success'} size="small">
              {error ? '⚠️ Error' : '📊 Analysis Result'}
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
