import React, { useCallback } from 'react';
import { makeStyles, tokens, Text, Button, Badge, Spinner, Tooltip } from '@fluentui/react-components';
import { MoneyRegular, CalculatorRegular, ArrowTrendingRegular, ScalesRegular, DocumentTextRegular, DismissRegular } from '@fluentui/react-icons';
import { MarkdownRenderer } from '../shared/MarkdownRenderer';
import { CopyButton } from '../shared/CopyButton';
import { useExcelTools } from '../../hooks/useExcelTools';

const FINANCE_SYSTEM_PROMPT = `You are an expert financial analyst working inside Microsoft Excel. You help with budgets, cash flow, financial ratios, break-even analysis, variance analysis, and financial modeling. Always provide Excel formulas, practical recommendations, and structured tables. Use standard financial terminology and explain calculations clearly.`;

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
  { id: 'budget', title: 'Budget Analysis', desc: 'Budget vs actual variance', icon: <MoneyRegular />,
    prompt: 'Analyze this budget data. Calculate: variance (amount and %), identify over/under-budget items, year-over-year comparison, and suggest areas for cost reduction. Provide a summary table with RAG status. Include Excel formulas for a budget tracking dashboard.' },
  { id: 'cashflow', title: 'Cash Flow Analysis', desc: 'Inflow/outflow patterns', icon: <ArrowTrendingRegular />,
    prompt: 'Analyze the cash flow data. Identify: net cash position, inflow/outflow patterns, burn rate, runway estimation, seasonal variations. Provide cash flow statement formulas and projections for the next 3 periods.' },
  { id: 'ratios', title: 'Financial Ratios', desc: 'Liquidity, profitability, efficiency', icon: <CalculatorRegular />,
    prompt: 'Calculate key financial ratios: liquidity (current ratio, quick ratio, cash ratio), profitability (gross margin, net margin, ROE, ROA, ROIC), efficiency (asset turnover, inventory turnover, DSO), leverage (debt-to-equity, interest coverage). Provide formulas, values, and interpretation.' },
  { id: 'breakeven', title: 'Break-Even Analysis', desc: 'Break-even point calculation', icon: <ScalesRegular />,
    prompt: 'Calculate the break-even point. Identify: fixed costs, variable costs per unit, contribution margin, break-even in units and revenue. Provide sensitivity analysis (what-if scenarios) and margin of safety. Include Excel formulas and a chart recommendation.' },
  { id: 'variance', title: 'Variance Analysis', desc: 'Actual vs planned differences', icon: <DocumentTextRegular />,
    prompt: 'Perform a variance analysis on this data. Calculate: favorable/unfavorable variances, percentage variances, root cause suggestions. Provide a waterfall chart recommendation and drill-down analysis.' },
  { id: 'profitability', title: 'Profitability Analysis', desc: 'Revenue, cost, margin analysis', icon: <MoneyRegular />,
    prompt: 'Analyze profitability from this data. Calculate: gross profit, operating profit, net profit, margin percentages, profit per product/segment (if applicable). Identify most/least profitable items. Suggest optimization strategies.' },
  { id: 'financial_model', title: 'Financial Model', desc: 'Build a model from data', icon: <CalculatorRegular />,
    prompt: 'Based on this data, suggest a financial model structure. Include: assumptions section, income statement projection, balance sheet items (if relevant), and a sensitivity/scenario analysis table. Provide all Excel formulas needed.' },
  { id: 'swot', title: 'SWOT from Data', desc: 'Data-driven SWOT analysis', icon: <DocumentTextRegular />,
    prompt: 'Perform a data-driven SWOT analysis based on the numbers in this spreadsheet. Identify: Strengths (strong metrics), Weaknesses (poor metrics), Opportunities (positive trends), Threats (negative trends). Support each point with specific data.' },
];

export const FinanceTools: React.FC = () => {
  const classes = useStyles();
  const { result, isLoading, error, runTool, clearResult, cancelStream } = useExcelTools();

  const handleTool = useCallback(async (toolId: string) => {
    const tool = TOOLS.find(t => t.id === toolId);
    if (!tool) return;
    await runTool(FINANCE_SYSTEM_PROMPT, tool.prompt, { includeSelection: true, includeWorkbook: true });
  }, [runTool]);

  return (
    <div className={classes.container}>
      <Text size={400} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <MoneyRegular /> Finance Tools
      </Text>
      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Select financial data, then choose an analysis</Text>

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
          <Spinner size="small" /><Text size={200}>Running financial analysis...</Text>
          <Button size="small" appearance="subtle" onClick={cancelStream}>Cancel</Button>
        </div>
      )}

      {(result || error) && (
        <div className={classes.resultArea}>
          <div className={classes.resultHeader}>
            <Badge appearance="outline" color={error ? 'danger' : 'success'} size="small">{error ? '⚠️ Error' : '💰 Financial Analysis'}</Badge>
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
