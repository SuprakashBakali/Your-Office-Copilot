import React, { useCallback } from 'react';
import { makeStyles, tokens, Text, Button, Badge, Spinner, Tooltip } from '@fluentui/react-components';
import { BroomRegular, CalendarRegular, TextSortAscendingRegular, CutRegular, SparkleRegular, DismissRegular, TextCaseTitleRegular } from '@fluentui/react-icons';
import { MarkdownRenderer } from '../shared/MarkdownRenderer';
import { CopyButton } from '../shared/CopyButton';
import { useExcelTools } from '../../hooks/useExcelTools';

const CLEANING_SYSTEM_PROMPT = `You are an expert data cleaning specialist working inside Microsoft Excel. You help users clean, standardize, and transform messy data. Always provide specific Excel formulas, step-by-step instructions, and explain the before/after for each fix. Prioritize non-destructive methods (new columns with formulas) over in-place modifications.`;

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
  { id: 'full-clean', title: 'AI Data Cleaning', desc: 'Full automated analysis', icon: <SparkleRegular />,
    prompt: 'Perform a comprehensive data quality audit. Check for: inconsistent formatting, mixed data types, extra/trailing spaces, special characters, missing values, obvious typos, inconsistent naming conventions, and formatting issues. For each problem: describe it, show examples from the data, and provide Excel formulas to fix it.' },
  { id: 'duplicates', title: 'Remove Duplicates', desc: 'Find and handle dupes', icon: <BroomRegular />,
    prompt: 'Analyze this data for duplicates. Report: exact duplicates (entire rows), partial duplicates (key column matches), and near-duplicates (fuzzy matches). For each: show the duplicate entries, suggest which to keep, and provide Excel formulas (COUNTIF, UNIQUE, Remove Duplicates steps).' },
  { id: 'dates', title: 'Standardize Dates', desc: 'Fix date format issues', icon: <CalendarRegular />,
    prompt: 'Identify all date-like values in this data and their current formats. Provide formulas to convert them all to a consistent ISO format (YYYY-MM-DD). Handle edge cases: text dates, US vs EU formats, 2-digit years, timestamps, relative dates ("yesterday", "last week").' },
  { id: 'split-merge', title: 'Split/Merge Columns', desc: 'Separate or combine fields', icon: <CutRegular />,
    prompt: 'Analyze this data for columns that should be split (e.g., "Full Name" → "First" + "Last", "Address" → components) or merged. Provide Excel formulas using: LEFT, RIGHT, MID, FIND, TEXTJOIN, CONCAT, Flash Fill patterns.' },
  { id: 'normalize', title: 'Normalize Text', desc: 'Case, spacing, encoding', icon: <TextCaseTitleRegular />,
    prompt: 'Normalize text data: fix case (PROPER, UPPER, LOWER as appropriate), trim whitespace (TRIM, CLEAN), remove non-printable characters, standardize abbreviations, fix encoding issues. Provide formulas for each fix.' },
  { id: 'inconsistent', title: 'Detect Inconsistencies', desc: 'Find conflicting values', icon: <TextSortAscendingRegular />,
    prompt: 'Detect inconsistent values: different spellings of the same thing (e.g., "USA" vs "United States" vs "US"), mixed units, inconsistent categories, and outlier text patterns. Suggest a standardization mapping and formulas (SWITCH, SUBSTITUTE, IFS) to harmonize the data.' },
];

export const CleaningTools: React.FC = () => {
  const classes = useStyles();
  const { result, isLoading, error, runTool, clearResult, cancelStream } = useExcelTools();

  const handleTool = useCallback(async (toolId: string) => {
    const tool = TOOLS.find(t => t.id === toolId);
    if (!tool) return;
    await runTool(CLEANING_SYSTEM_PROMPT, tool.prompt, { includeSelection: true, includeFormulas: true });
  }, [runTool]);

  return (
    <div className={classes.container}>
      <Text size={400} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <BroomRegular /> Data Cleaning
      </Text>
      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Select messy data, then choose a cleaning tool</Text>

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
          <Spinner size="small" /><Text size={200}>Scanning data for issues...</Text>
          <Button size="small" appearance="subtle" onClick={cancelStream}>Cancel</Button>
        </div>
      )}

      {(result || error) && (
        <div className={classes.resultArea}>
          <div className={classes.resultHeader}>
            <Badge appearance="outline" color={error ? 'danger' : 'success'} size="small">{error ? '⚠️ Error' : '🧹 Cleaning Results'}</Badge>
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
