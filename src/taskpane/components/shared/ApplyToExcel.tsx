import React, { useState } from 'react';
import { Button, Spinner, Text, makeStyles, tokens } from '@fluentui/react-components';
import { ArrowDownloadRegular, CheckmarkCircleRegular, ErrorCircleRegular } from '@fluentui/react-icons';
import { ExcelService } from '../../services/office/ExcelService';

const useStyles = makeStyles({
  bar: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '6px',
    alignItems: 'center',
    padding: '8px',
    borderRadius: '8px',
    backgroundColor: tokens.colorBrandBackground2,
    border: `1px solid ${tokens.colorBrandStroke2}`,
    marginTop: '8px',
  },
  status: {
    fontSize: tokens.fontSizeBase100,
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },
  success: { color: tokens.colorPaletteGreenForeground1 },
  error: { color: tokens.colorPaletteRedForeground1 },
});

/**
 * Extract the first Excel formula (=...) from AI-generated markdown text.
 * Looks inside ```excel, ```xlsx, and plain code blocks, or raw = lines.
 */
export function extractFirstFormula(text: string): string | null {
  // Match code blocks: ```excel\n=...\n```
  const codeBlockMatch = text.match(/```(?:excel|xlsx|formula|vba)?\s*\n?(=[\s\S]*?)\n?```/i);
  if (codeBlockMatch) return codeBlockMatch[1].trim().split('\n')[0].trim();

  // Match inline code: `=FORMULA(...)`
  const inlineMatch = text.match(/`(=[^`]+)`/);
  if (inlineMatch) return inlineMatch[1].trim();

  // Match bare formula lines: lines starting with =
  const bareMatch = text.match(/^(=[A-Z(][\s\S]*?)$/m);
  if (bareMatch) return bareMatch[1].trim().split('\n')[0].trim();

  return null;
}

interface ApplyToExcelProps {
  resultText: string;
  /** If provided, shows an "Apply Formula" button and inserts into active cell */
  showApplyFormula?: boolean;
  /** If provided, shows a custom action button */
  customAction?: {
    label: string;
    icon?: React.ReactElement;
    onApply: () => Promise<void>;
  };
}

export const ApplyToExcel: React.FC<ApplyToExcelProps> = ({
  resultText,
  showApplyFormula = true,
  customAction,
}) => {
  const classes = useStyles();
  const [status, setStatus] = useState<'idle' | 'applying' | 'success' | 'error'>('idle');
  const [statusMsg, setStatusMsg] = useState('');

  const formula = showApplyFormula ? extractFirstFormula(resultText) : null;

  const applyFormula = async () => {
    if (!formula) return;
    setStatus('applying');
    setStatusMsg('');
    try {
      // Get currently selected cell address
      const selection = await ExcelService.getSelectedRange();
      // Use first cell of the selection
      const firstCell = selection.address.split(':')[0].replace(/^[^!]+!/, '');
      await ExcelService.insertFormula(firstCell, formula);
      setStatus('success');
      setStatusMsg(`Applied to ${firstCell}`);
      setTimeout(() => setStatus('idle'), 3000);
    } catch (e) {
      setStatus('error');
      setStatusMsg((e as Error).message);
    }
  };

  const handleCustom = async () => {
    if (!customAction) return;
    setStatus('applying');
    setStatusMsg('');
    try {
      await customAction.onApply();
      setStatus('success');
      setStatusMsg('Applied successfully');
      setTimeout(() => setStatus('idle'), 3000);
    } catch (e) {
      setStatus('error');
      setStatusMsg((e as Error).message);
    }
  };

  if (!formula && !customAction) return null;

  return (
    <div className={classes.bar}>
      {formula && (
        <Button
          appearance="primary"
          size="small"
          icon={status === 'applying' ? <Spinner size="extra-tiny" /> : <ArrowDownloadRegular />}
          onClick={applyFormula}
          disabled={status === 'applying'}
        >
          Apply Formula to Active Cell
        </Button>
      )}
      {customAction && (
        <Button
          appearance="primary"
          size="small"
          icon={status === 'applying' ? <Spinner size="extra-tiny" /> : (customAction.icon || <ArrowDownloadRegular />)}
          onClick={handleCustom}
          disabled={status === 'applying'}
        >
          {customAction.label}
        </Button>
      )}
      {status === 'success' && (
        <Text className={`${classes.status} ${classes.success}`}>
          <CheckmarkCircleRegular /> {statusMsg}
        </Text>
      )}
      {status === 'error' && (
        <Text className={`${classes.status} ${classes.error}`}>
          <ErrorCircleRegular /> {statusMsg}
        </Text>
      )}
    </div>
  );
};
