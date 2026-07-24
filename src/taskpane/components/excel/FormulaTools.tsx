import React, { useState, useCallback } from 'react';
import { makeStyles, tokens, Card, CardHeader, Text, Button, Textarea, Spinner, Tooltip, Badge } from '@fluentui/react-components';
import {
  CalculatorRegular, DocumentSearchRegular, WrenchRegular, ArrowSyncRegular,
  LightbulbRegular, TextGrammarWandRegular, ArrowDownRegular, TableSimpleRegular,
  DismissRegular, CheckmarkRegular, SearchRegular, FilterRegular, CodeRegular, TableRegular, DatabaseRegular,
} from '@fluentui/react-icons';
import { MarkdownRenderer } from '../shared/MarkdownRenderer';
import { CopyButton } from '../shared/CopyButton';
import { useExcelTools } from '../../hooks/useExcelTools';

const FORMULA_SYSTEM_PROMPT = `You are an expert Excel formula assistant. You help users with:
- Generating Excel formulas from plain English descriptions
- Explaining existing formulas step-by-step
- Finding and fixing formula errors
- Suggesting optimized alternatives using modern Excel functions (XLOOKUP, FILTER, UNIQUE, SORT, LET, LAMBDA)
- Converting formulas to/from plain English

Always provide the formula in a code block with the \`excel\` language identifier.
If explaining, break down each component clearly.
If suggesting alternatives, explain why the alternative is better.`;

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    padding: '16px',
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
    transition: 'border-color 0.2s ease, box-shadow 0.2s ease',
    '&:hover': {
      border: `1px solid ${tokens.colorBrandStroke1}`,
      boxShadow: `0 0 0 1px ${tokens.colorBrandStroke1}`,
    },
  },
  sectionHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    color: tokens.colorBrandForeground1,
  },
  textarea: {
    minHeight: '60px',
  },
  resultArea: {
    padding: '14px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: '10px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    position: 'relative',
  },
  resultHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '8px',
  },
  resultActions: {
    display: 'flex',
    gap: '4px',
    marginTop: '12px',
  },
  quickActions: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: '8px',
  },
  quickCard: {
    cursor: 'pointer',
    padding: '10px',
    borderRadius: '8px',
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    transition: 'all 0.2s ease',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      border: `1px solid ${tokens.colorBrandStroke1}`,
      transform: 'translateY(-1px)',
    },
  },
});

export const FormulaTools: React.FC = () => {
  const classes = useStyles();
  const [description, setDescription] = useState('');
  const { result, isLoading, error, runTool, clearResult, cancelStream } = useExcelTools();

  const generateFormula = useCallback(async () => {
    if (!description.trim()) return;
    await runTool(
      FORMULA_SYSTEM_PROMPT,
      `Generate an Excel formula that does the following: ${description}\n\nProvide the formula, explain how it works, and give usage instructions.`,
      { includeSelection: true, includeFormulas: true }
    );
  }, [description, runTool]);

  const explainFormula = useCallback(async () => {
    await runTool(
      FORMULA_SYSTEM_PROMPT,
      'Explain the formula(s) in the selected cells step-by-step. Break down each function, what it does, and how the parts work together. Also suggest if there are simpler or more modern alternatives.',
      { includeSelection: true, includeFormulas: true }
    );
  }, [runTool]);

  const fixFormula = useCallback(async () => {
    await runTool(
      FORMULA_SYSTEM_PROMPT,
      'The selected cells contain formula(s) that may have errors. Diagnose any issues, explain why they might fail, and provide corrected versions. Check for: circular references, wrong ranges, missing parentheses, data type mismatches, #N/A, #REF!, #VALUE!, #DIV/0! errors.',
      { includeSelection: true, includeFormulas: true }
    );
  }, [runTool]);

  const optimizeFormula = useCallback(async () => {
    await runTool(
      FORMULA_SYSTEM_PROMPT,
      'Review and optimize the formula(s) in the selected cells. Consider: performance (INDEX/MATCH or XLOOKUP over VLOOKUP), readability, error handling (IFERROR/IFNA), and modern Excel functions (FILTER, UNIQUE, SORT, SORTBY, LET, LAMBDA, MAP, REDUCE). Show before/after.',
      { includeSelection: true, includeFormulas: true }
    );
  }, [runTool]);

  const formulaToEnglish = useCallback(async () => {
    await runTool(
      FORMULA_SYSTEM_PROMPT,
      'Convert the formula(s) in the selected cells to plain English. Explain what they calculate in simple, non-technical language that anyone could understand.',
      { includeSelection: true, includeFormulas: true }
    );
  }, [runTool]);

  const generateXlookup = useCallback(async () => {
    await runTool(FORMULA_SYSTEM_PROMPT, 'I need an XLOOKUP formula based on this context. Please provide the exact XLOOKUP formula and explain its arguments.', { includeSelection: true });
  }, [runTool]);

  const generateFilterSortUnique = useCallback(async () => {
    await runTool(FORMULA_SYSTEM_PROMPT, 'I need a formula using modern dynamic arrays like FILTER, SORT, and UNIQUE based on this context.', { includeSelection: true });
  }, [runTool]);

  const generateLetLambda = useCallback(async () => {
    await runTool(FORMULA_SYSTEM_PROMPT, 'I need a complex formula optimized using the LET function, or a custom LAMBDA function based on this context.', { includeSelection: true });
  }, [runTool]);

  const generateDynamicArray = useCallback(async () => {
    await runTool(FORMULA_SYSTEM_PROMPT, 'I need a dynamic array formula (spilling behavior) to solve the problem described in this context.', { includeSelection: true });
  }, [runTool]);

  const generateDax = useCallback(async () => {
    await runTool(FORMULA_SYSTEM_PROMPT, 'Write a DAX measure or calculated column formula for this data model context.', { includeSelection: true });
  }, [runTool]);

  return (
    <div className={classes.container}>
      <Text size={400} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <CalculatorRegular /> Formula Tools
      </Text>

      {/* Generate Formula */}
      <div className={classes.section}>
        <div className={classes.sectionHeader}>
          <DocumentSearchRegular />
          <Text weight="semibold">Generate Formula</Text>
        </div>
        <Textarea
          className={classes.textarea}
          placeholder="Describe what you want to calculate... e.g., 'Sum all values in column B where column A says Sales'"
          value={description}
          onChange={(_, d) => setDescription(d.value)}
          resize="vertical"
        />
        <Button
          appearance="primary"
          onClick={generateFormula}
          disabled={!description.trim() || isLoading}
          icon={isLoading ? <Spinner size="tiny" /> : <LightbulbRegular />}
        >
          {isLoading ? 'Generating...' : 'Generate'}
        </Button>
      </div>

      {/* Quick Actions */}
      <div className={classes.quickActions}>
        <div className={classes.quickCard} onClick={explainFormula}>
          <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <TextGrammarWandRegular /> Explain Formula
          </Text>
          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Break down selected formula</Text>
        </div>
        <div className={classes.quickCard} onClick={fixFormula}>
          <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <WrenchRegular /> Fix Errors
          </Text>
          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Debug formula issues</Text>
        </div>
        <div className={classes.quickCard} onClick={optimizeFormula}>
          <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <ArrowSyncRegular /> Optimize
          </Text>
          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Improve performance</Text>
        </div>
        <div className={classes.quickCard} onClick={formulaToEnglish}>
          <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <ArrowDownRegular /> To English
          </Text>
          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Formula → plain language</Text>
        </div>
        <div className={classes.quickCard} onClick={generateXlookup}>
          <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <SearchRegular /> XLOOKUP
          </Text>
          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>XLOOKUP specialist</Text>
        </div>
        <div className={classes.quickCard} onClick={generateFilterSortUnique}>
          <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <FilterRegular /> FILTER / SORT
          </Text>
          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Modern array tools</Text>
        </div>
        <div className={classes.quickCard} onClick={generateLetLambda}>
          <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <CodeRegular /> LET & LAMBDA
          </Text>
          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Advanced functions</Text>
        </div>
        <div className={classes.quickCard} onClick={generateDynamicArray}>
          <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <TableRegular /> Dynamic Arrays
          </Text>
          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Spilling formulas</Text>
        </div>
        <div className={classes.quickCard} onClick={generateDax}>
          <Text size={200} weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <DatabaseRegular /> DAX Writer
          </Text>
          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Power Pivot DAX</Text>
        </div>
      </div>

      {/* Result */}
      {(result || error) && (
        <div className={classes.resultArea}>
          <div className={classes.resultHeader}>
            <Badge appearance="outline" color={error ? 'danger' : 'success'} size="small">
              {error ? '⚠️ Error' : '✅ Result'}
            </Badge>
            <div style={{ display: 'flex', gap: '2px' }}>
              <CopyButton text={result} />
              {isLoading && (
                <Tooltip content="Stop generating" relationship="label">
                  <Button appearance="transparent" size="small" icon={<DismissRegular />} onClick={cancelStream} />
                </Tooltip>
              )}
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
