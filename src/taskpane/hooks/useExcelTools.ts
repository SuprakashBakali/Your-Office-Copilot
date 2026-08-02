import { useState, useCallback, useRef } from 'react';
import { useAI } from './useAI';
import { ExcelService } from '../services/office/ExcelService';
import { CellData } from '../types';
import { useSettings } from './useSettings';

export interface ToolResult {
  content: string;
  isLoading: boolean;
  error: string | null;
}

/**
 * Hook that encapsulates the pattern: read Excel data → build prompt → call AI → return result.
 * Used by all Excel tool components.
 */
export function useExcelTools() {
  const [result, setResult] = useState<string>('');
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const ai = useAI();
  const { settings } = useSettings();
  // H3: keep settings in a ref so runTool's useCallback dep array only needs
  // [ai] — otherwise a new settings object ref on every render invalidates
  // runTool and cascades stale-closure issues to every child useCallback.
  const settingsRef = useRef(settings);
  settingsRef.current = settings;

  /**
   * Run a tool: reads Excel context, combines with a prompt, sends to AI
   */
  const runTool = useCallback(async (
    systemPrompt: string,
    userPrompt: string,
    options?: {
      includeSelection?: boolean;
      includeWorkbook?: boolean;
      includeFormulas?: boolean;
      customContext?: string;
    }
  ): Promise<string> => {
    setIsLoading(true);
    setError(null);
    setResult('');

    try {
      let contextStr = '';

      if (options?.customContext) {
        contextStr = options.customContext;
      } else {
        try {
          if (options?.includeSelection !== false) {
            const data = await ExcelService.getSelectedRange();
            contextStr += `\n\nSelected Range: ${data.address} (Sheet: ${data.sheetName})\n`;
            contextStr += `Dimensions: ${data.rowCount} rows × ${data.columnCount} columns\n`;
            // M5: cap rows to 100 AND columns to 50 — an unbounded column slice
            // (e.g. 1000-col CSV) would bloat the prompt to hundreds of KB.
            // Also respect maxContextCells: derive max rows from the cell budget
            // (e.g. budget=500 ÷ 50 cols = 10 rows).
            const maxCells = settingsRef.current.maxContextCells || 5000;
            const maxCols = Math.min(data.columnCount, 50);
            const maxRows = Math.min(data.rowCount, 100, Math.ceil(maxCells / Math.max(maxCols, 1)));
            contextStr += `Data:\n${JSON.stringify(data.values.slice(0, maxRows).map(row => row.slice(0, maxCols)))}\n`;

            if (options?.includeFormulas && data.formulas) {
              const formulaCells: string[] = [];
              data.formulas.forEach((row, ri) => {
                row.forEach((cell, ci) => {
                  if (typeof cell === 'string' && cell.startsWith('=')) {
                    formulaCells.push(`  ${data.address.split('!')[0]}!R${ri+1}C${ci+1}: ${cell}`);
                  }
                });
              });
              if (formulaCells.length > 0) {
                contextStr += `\nFormulas found:\n${formulaCells.join('\n')}\n`;
              }
            }
          }

          if (options?.includeWorkbook) {
            try {
              const wbInfo = await ExcelService.getWorkbookInfo();
              contextStr += `\nWorkbook: ${wbInfo.name}\n`;
              contextStr += `Sheets: ${wbInfo.sheets.map(s => `${s.name}${s.isVisible ? '' : ' (hidden)'}`).join(', ')}\n`;
              contextStr += `Active Sheet: ${wbInfo.activeSheet}\n`;
              if (wbInfo.namedRanges && wbInfo.namedRanges.length > 0) {
                contextStr += `Named Ranges: ${wbInfo.namedRanges.join(', ')}\n`;
              }
            } catch {
              // Workbook info not critical
            }
          }
        } catch (officeErr) {
          contextStr = `\n[Note: Could not read Excel data — ${(officeErr as Error).message}. Please describe your data in the prompt.]\n`;
        }
      }

      const fullSystemPrompt = `${systemPrompt}\n\nIMPORTANT: Format your response using Markdown. Use tables, bullet points, code blocks, and headings for clarity.${contextStr ? `\n\nCurrent Excel context:${contextStr}` : ''}`;

      const messages = [
        { id: 'sys', role: 'system' as const, content: fullSystemPrompt, timestamp: 0 },
        { id: 'usr', role: 'user' as const, content: userPrompt, timestamp: Date.now() },
      ];

      let response = '';

      if (settingsRef.current.streamResponses) {
        response = (await ai.sendMessageStream(messages, {}, (chunk) => {
          setResult(prev => prev + chunk);
        })).text;
      } else {
        response = await ai.sendMessage(messages);
        setResult(response);
      }

      setIsLoading(false);
      return response;
    } catch (err) {
      const msg = (err as Error).message;
      setError(msg);
      setIsLoading(false);
      throw err;
    }
  // H3: only [ai] in deps — settings is accessed via settingsRef.current
  }, [ai]);

  /**
   * Get the currently selected data from Excel
   */
  const getSelection = useCallback(async (): Promise<CellData | null> => {
    try {
      return await ExcelService.getSelectedRange();
    } catch {
      return null;
    }
  }, []);

  /**
   * Insert a formula into the active cell
   */
  const applyFormula = useCallback(async (address: string, formula: string) => {
    try {
      await ExcelService.insertFormula(address, formula);
      return true;
    } catch {
      return false;
    }
  }, []);

  /**
   * Write values to a range
   */
  const applyValues = useCallback(async (address: string, values: any[][]) => {
    try {
      await ExcelService.writeToRange(address, values);
      return true;
    } catch {
      return false;
    }
  }, []);

  /**
   * Create a chart
   */
  const createChart = useCallback(async (type: string, dataRange: string, title: string) => {
    try {
      await ExcelService.createChart(type, dataRange, title);
      return true;
    } catch {
      return false;
    }
  }, []);

  const clearResult = useCallback(() => {
    setResult('');
    setError(null);
  }, []);

  return {
    result,
    isLoading,
    error,
    runTool,
    getSelection,
    applyFormula,
    applyValues,
    createChart,
    clearResult,
    cancelStream: ai.cancelStream,
  };
}
