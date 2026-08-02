import { CellData, WorkbookInfo, SheetInfo } from '../../types';

export class ExcelService {
  static async getSelectedRange(): Promise<CellData> {
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "values", "formulas", "numberFormat", "rowCount", "columnCount"]);
      const worksheet = range.worksheet;
      worksheet.load("name");
      await context.sync();
      
      return {
        address: range.address,
        values: range.values,
        formulas: range.formulas,
        numberFormat: range.numberFormat,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        sheetName: worksheet.name
      };
    });
  }

  static async getUsedRange(): Promise<CellData> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getUsedRange();
      range.load(["address", "values", "formulas", "numberFormat", "rowCount", "columnCount"]);
      worksheet.load("name");
      await context.sync();
      
      return {
        address: range.address,
        values: range.values,
        formulas: range.formulas,
        numberFormat: range.numberFormat,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        sheetName: worksheet.name
      };
    });
  }

  static async getSheetNames(): Promise<string[]> {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      return sheets.items.map(sheet => sheet.name);
    });
  }

  static async getWorkbookInfo(): Promise<WorkbookInfo> {
    return Excel.run(async (context) => {
      const workbook = context.workbook;
      const worksheets = workbook.worksheets;
      const activeWorksheet = worksheets.getActiveWorksheet();
      const namedItems = workbook.names;
      
      workbook.load("name");
      worksheets.load("items/name, items/visibility");
      activeWorksheet.load("name");
      namedItems.load("items/name");
      
      await context.sync();
      
      const sheetInfos: SheetInfo[] = [];
      for (const sheet of worksheets.items) {
        let rowCount = 0;
        let columnCount = 0;
        let usedRangeAddress = "";
        
        try {
          const usedRange = sheet.getUsedRange();
          usedRange.load(["rowCount", "columnCount", "address"]);
          await context.sync();
          rowCount = usedRange.rowCount;
          columnCount = usedRange.columnCount;
          usedRangeAddress = usedRange.address;
        } catch (e) {
          // Empty sheet or error
        }
        
        sheetInfos.push({
          name: sheet.name,
          isVisible: sheet.visibility === Excel.SheetVisibility.visible,
          rowCount,
          columnCount,
          usedRangeAddress
        });
      }
      
      return {
        name: workbook.name || 'Workbook',
        sheets: sheetInfos,
        activeSheet: activeWorksheet.name,
        namedRanges: namedItems.items.map(item => item.name)
      };
    });
  }

  static async getSheetData(sheetName: string): Promise<CellData> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(sheetName);
      const range = worksheet.getUsedRange();
      range.load(["address", "values", "formulas", "numberFormat", "rowCount", "columnCount"]);
      worksheet.load("name");
      await context.sync();
      
      return {
        address: range.address,
        values: range.values,
        formulas: range.formulas,
        numberFormat: range.numberFormat,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        sheetName: worksheet.name
      };
    });
  }

  static async getNamedRanges(): Promise<string[]> {
    return Excel.run(async (context) => {
      const namedItems = context.workbook.names;
      namedItems.load("items/name");
      await context.sync();
      return namedItems.items.map(item => item.name);
    });
  }

  static async getCellFormulas(address: string): Promise<string[][]> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.load("formulas");
      await context.sync();
      return range.formulas;
    });
  }

  static async writeToRange(address: string, values: any[][]): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.values = values;
      await context.sync();
    });
  }

  static async insertFormula(address: string, formula: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.formulas = [[formula]];
      await context.sync();
    });
  }

  static async applyConditionalFormatting(address: string, rules: any): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      if (Office.context.requirements.isSetSupported('ExcelApi', '1.6')) {
        // Parse the rules parameter to determine which conditional format type to apply.
        // If no rules specified, default to color scale.
        const ruleType = typeof rules === 'string' ? rules.toLowerCase() : (rules?.type || 'colorscale');

        if (ruleType === 'colorscale' || ruleType === 'color_scale') {
          const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
          cf.colorScale.criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
            midpoint: { type: Excel.ConditionalFormatColorCriterionType.percentile, formula: "50", color: "#FFEB84" },
            maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" }
          };
        } else if (ruleType === 'databar' || ruleType === 'data_bar') {
          const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
          (cf.dataBar as any).barColor = rules?.color || "#76B900";
        } else if (ruleType === 'iconset' || ruleType === 'icon_set') {
          const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
          cf.iconSet.style = Excel.IconSet.threeTrafficLights1;
        } else {
          // Default: color scale
          const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
          cf.colorScale.criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
            midpoint: { type: Excel.ConditionalFormatColorCriterionType.percentile, formula: "50", color: "#FFEB84" },
            maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" }
          };
        }
      }
      await context.sync();
    });
  }

  static async createNamedRange(name: string, address: string): Promise<void> {
    return Excel.run(async (context) => {
      context.workbook.names.add(name, address);
      await context.sync();
    });
  }

  static async createChart(type: string, dataRange: string, title: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(dataRange);
      let chartType = Excel.ChartType.columnClustered;
      const t = type.toLowerCase();
      if (t.includes('pie')) chartType = Excel.ChartType.pie;
      else if (t.includes('line')) chartType = Excel.ChartType.line;
      else if (t.includes('bar')) chartType = Excel.ChartType.barClustered;
      else if (t.includes('area')) chartType = Excel.ChartType.area;
      else if (t.includes('scatter')) chartType = Excel.ChartType.xyscatter;
      const chart = worksheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);
      chart.title.text = title;
      await context.sync();
    });
  }

  static async deleteChart(chartName: string): Promise<void> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      if (chartName.toLowerCase() === 'all') {
        const charts = sheet.charts;
        charts.load("items");
        await context.sync();
        charts.items.forEach(c => c.delete());
      } else {
        const chart = sheet.charts.getItemOrNullObject(chartName);
        await context.sync();
        if (!chart.isNullObject) {
          chart.delete();
        }
      }
      await context.sync();
    });
  }

  static async clearRange(address: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.clear();
      await context.sync();
    });
  }

  static async formatRange(
    address: string,
    opts: { 
      bold?: boolean; italic?: boolean; backgroundColor?: string; fontColor?: string; fontSize?: number;
      wrapText?: boolean; horizontalAlignment?: "Center" | "Left" | "Right" | "Justify" | "General"; 
      verticalAlignment?: "Center" | "Top" | "Bottom" | "Justify"; numberFormat?: string 
    }
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      if (opts.bold !== undefined) range.format.font.bold = opts.bold;
      if (opts.italic !== undefined) range.format.font.italic = opts.italic;
      if (opts.backgroundColor) range.format.fill.color = opts.backgroundColor;
      if (opts.fontColor) range.format.font.color = opts.fontColor;
      if (opts.fontSize) range.format.font.size = opts.fontSize;
      if (opts.wrapText !== undefined) range.format.wrapText = opts.wrapText;
      if (opts.horizontalAlignment) range.format.horizontalAlignment = opts.horizontalAlignment as any;
      if (opts.verticalAlignment) range.format.verticalAlignment = opts.verticalAlignment as any;
      if (opts.numberFormat) {
        range.numberFormat = [[opts.numberFormat]];
      }
      await context.sync();
    });
  }

  /**
   * Create a PivotTable from a source range.
   * Requires ExcelApi 1.8+ (available in Microsoft 365).
   */
  static async createPivotTable(
    sourceRange: string,
    targetCell: string,
    rowField: string,
    valueField: string,
    pivotName = 'AICopilotPivot'
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const src = worksheet.getRange(sourceRange);
      const dest = worksheet.getRange(targetCell);

      // Remove existing pivot with same name if present
      try {
        const existing = worksheet.pivotTables.getItem(pivotName);
        existing.delete();
        await context.sync();
      } catch { /* doesn't exist yet */ }

      const pivotTable = worksheet.pivotTables.add(pivotName, src, dest);
      await context.sync();

      // Add row field
      try {
        const rowHierarchy = pivotTable.hierarchies.getItem(rowField);
        pivotTable.rowHierarchies.add(rowHierarchy);
        await context.sync();
      } catch {
        // field name might differ; load all hierarchies for debug
        pivotTable.hierarchies.load('items/name');
        await context.sync();
        const available = pivotTable.hierarchies.items.map(h => h.name).join(', ');
        throw new Error(`Row field "${rowField}" not found. Available fields: ${available}`);
      }

      // Add value field
      try {
        const valueHierarchy = pivotTable.hierarchies.getItem(valueField);
        pivotTable.dataHierarchies.add(valueHierarchy);
        await context.sync();
      } catch {
        pivotTable.hierarchies.load('items/name');
        await context.sync();
        const available = pivotTable.hierarchies.items.map(h => h.name).join(', ');
        throw new Error(`Value field "${valueField}" not found. Available fields: ${available}`);
      }
    });
  }

  static async addSheet(name: string): Promise<void> {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.add(name);
      await context.sync();
    });
  }

  static async deleteSheet(name: string): Promise<void> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(name);
      sheet.delete();
      await context.sync();
    });
  }

  static async insertRange(address: string, shiftDirection: 'Down' | 'Right'): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.insert(shiftDirection === 'Down' ? Excel.InsertShiftDirection.down : Excel.InsertShiftDirection.right);
      await context.sync();
    });
  }

  static async deleteRange(address: string, shiftDirection: 'Up' | 'Left'): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.delete(shiftDirection === 'Up' ? Excel.DeleteShiftDirection.up : Excel.DeleteShiftDirection.left);
      await context.sync();
    });
  }

  static async mergeCells(address: string, mergeAcross: boolean = false): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.merge(mergeAcross);
      await context.sync();
    });
  }

  static async createTable(address: string, hasHeaders: boolean, name?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      const table = worksheet.tables.add(range, hasHeaders);
      if (name) table.name = name;
      await context.sync();
    });
  }

  static async sortRange(address: string, columnIndex: number, ascending: boolean = true): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.sort.apply([{ key: columnIndex, ascending }]);
      await context.sync();
    });
  }

  static async findAndReplace(address: string, findText: string, replaceText: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.replaceAll(findText, replaceText, { completeMatch: false, matchCase: false });
      await context.sync();
    });
  }

  static async addDataValidation(address: string, sourceList: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: sourceList
        }
      };
      await context.sync();
    });
  }

  static async addConditionalFormatting(address: string, type: 'colorScale' | 'dataBar'): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      if (type === 'colorScale') {
        range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      } else if (type === 'dataBar') {
        range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
      }
      await context.sync();
    });
  }

  // --- 12 New Data Cleaning & Analyzing Features ---

  static async removeDuplicates(address: string, columns?: number[]): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      if (columns && columns.length > 0) {
        range.removeDuplicates(columns, true);
      } else {
        range.removeDuplicates([], true);
      }
      await context.sync();
    });
  }

  static async trimWhitespace(address: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.load('values');
      await context.sync();
      const newValues = range.values.map(row => 
        row.map(cell => typeof cell === 'string' ? cell.trim() : cell)
      );
      range.values = newValues;
      await context.sync();
    });
  }

  static async changeCase(address: string, type: 'upper' | 'lower' | 'proper'): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.load('values');
      await context.sync();
      
      const toProper = (str: string) => str.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.slice(1).toLowerCase());
      
      const newValues = range.values.map(row => 
        row.map(cell => {
          if (typeof cell === 'string') {
            if (type === 'upper') return cell.toUpperCase();
            if (type === 'lower') return cell.toLowerCase();
            if (type === 'proper') return toProper(cell);
          }
          return cell;
        })
      );
      range.values = newValues;
      await context.sync();
    });
  }

  static async removeBlankRows(address: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.load(['values', 'rowCount', 'address']);
      await context.sync();
      
      const rowsToDelete: number[] = [];
      for (let r = 0; r < range.rowCount; r++) {
        const rowData = range.values[r];
        const isBlank = rowData.every(val => val === '' || val === null || val === undefined);
        if (isBlank) rowsToDelete.push(r);
      }
      
      // Delete from bottom to top to avoid shifting indexes
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        const rowRange = range.getRow(rowsToDelete[i]);
        rowRange.delete(Excel.DeleteShiftDirection.up);
      }
      await context.sync();
    });
  }

  static async applyFilter(address: string, columnIndex: number, criteria: string[]): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      worksheet.autoFilter.apply(range, columnIndex, {
        filterOn: Excel.FilterOn.values,
        values: criteria
      });
      await context.sync();
    });
  }

  static async clearFilter(): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.autoFilter.clearCriteria();
      await context.sync();
    });
  }

  static async groupData(address: string, byRows: boolean = true): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.group(byRows ? Excel.GroupOption.byRows : Excel.GroupOption.byColumns);
      await context.sync();
    });
  }

  static async ungroupData(address: string, byRows: boolean = true): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      range.ungroup(byRows ? Excel.GroupOption.byRows : Excel.GroupOption.byColumns);
      await context.sync();
    });
  }

  static async addSparklines(address: string, sourceAddress: string, type: 'line' | 'column' = 'line'): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const destRange = worksheet.getRange(address);
      const srcRange = worksheet.getRange(sourceAddress);
      (worksheet as any).sparklineGroups.add(
        type === 'line' ? 'Line' : 'Column',
        destRange,
        srcRange
      );
      await context.sync();
    });
  }

  static async formatChart(chartName: string, opts: { title?: string, showDataLabels?: boolean, legendPosition?: Excel.ChartLegendPosition }): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const chart = worksheet.charts.getItem(chartName);
      if (opts.title) {
        chart.title.text = opts.title;
        chart.title.visible = true;
      }
      if (opts.showDataLabels !== undefined) {
        chart.dataLabels.showValue = opts.showDataLabels;
      }
      if (opts.legendPosition) {
        chart.legend.position = opts.legendPosition;
        chart.legend.visible = true;
      }
      await context.sync();
    });
  }

  static async highlightDuplicates(address: string, color: string = 'pink'): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);
      conditionalFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.duplicateValues };
      conditionalFormat.preset.format.fill.color = color;
      conditionalFormat.preset.format.font.color = '#9C0006';
      await context.sync();
    });
  }

  static async highlightTopBottom(address: string, type: 'top' | 'bottom', count: number = 10, color: string = 'lightgreen'): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(address);
      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
      conditionalFormat.topBottom.rule = { rank: count, type: type === 'top' ? Excel.ConditionalTopBottomCriterionType.topItems : Excel.ConditionalTopBottomCriterionType.bottomItems };
      conditionalFormat.topBottom.format.fill.color = color;
      await context.sync();
    });
  }


  /**
   * Format 2D cell values cleanly for AI context, trimming trailing empty rows
   * and capping to maxCells to avoid multi-megabyte payloads when selecting entire columns.
   */
  private static formatCellValuesForAI(
    values: any[][],
    maxCells: number,
  ): { formatted: string; isTruncated: boolean; rowsShown: number; totalRows: number } {
    if (!values || !Array.isArray(values) || values.length === 0) {
      return { formatted: '[]', isTruncated: false, rowsShown: 0, totalRows: 0 };
    }

    // 1. Trim trailing completely empty rows (common when selecting columns like D:D)
    let lastNonEmptyRow = values.length - 1;
    while (lastNonEmptyRow >= 0) {
      const row = values[lastNonEmptyRow];
      const hasData =
        row &&
        Array.isArray(row) &&
        row.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== '');
      if (hasData) break;
      lastNonEmptyRow--;
    }

    const trimmedRows = lastNonEmptyRow >= 0 ? values.slice(0, lastNonEmptyRow + 1) : [];
    if (trimmedRows.length === 0) {
      return { formatted: '[]', isTruncated: false, rowsShown: 0, totalRows: values.length };
    }

    // 2. Cap to maxCells to prevent token overflow / HTTP 413 Payload Too Large
    const colCount = Math.max(1, trimmedRows[0]?.length || 1);
    const maxRowsAllowed = Math.max(10, Math.floor(maxCells / colCount));
    const rowsShown = Math.min(trimmedRows.length, maxRowsAllowed);
    const isTruncated = trimmedRows.length > maxRowsAllowed;

    const sliced = trimmedRows.slice(0, rowsShown);
    return {
      formatted: JSON.stringify(sliced),
      isTruncated,
      rowsShown,
      totalRows: trimmedRows.length,
    };
  }

  static async getContextForAI(maxCells: number = 1000): Promise<string> {
    try {
      const data = await this.getSelectedRange();
      let target = data;
      let label = 'Selected Range';
      if (data.rowCount * data.columnCount === 1 && !data.values[0][0]) {
        target = await this.getUsedRange();
        label = 'Used Range';
      }

      const { formatted, isTruncated, rowsShown, totalRows } = this.formatCellValuesForAI(
        target.values,
        maxCells,
      );
      let res = `Active Sheet: ${target.sheetName}\n${label}: ${target.address}\nData:\n${formatted}`;
      if (isTruncated) {
        res += `\n(Note: showing first ${rowsShown} of ${totalRows} non-empty rows in range to conserve tokens)`;
      }
      return res;
    } catch (e) {
      return `Unable to read Excel context: ${(e as Error).message}`;
    }
  }

  // ── From office-agents: eval_officejs escape hatch ────────────────────────
  /**
   * Execute arbitrary Office.js code inside Excel.run.
   * Returns the value returned by the code, or null.
   * The code string runs as: `async (context) => { <code> }`
   */
  static async evalOfficeJs(code: string): Promise<any> {
    return Excel.run(async (context) => {
      // Build an async function with `context` in scope and evaluate it
      // eslint-disable-next-line no-new-func
      const fn = new Function('context', 'Excel', `return (async () => { ${code} })()`);
      const result = await fn(context, Excel);
      return result ?? null;
    });
  }

  // ── From office-agents: screenshot_range ─────────────────────────────────
  /**
   * Capture a cell range as a base64 PNG with row/column headers composited
   * onto the image. Returns null if the API is not supported.
   */
  static async screenshotRange(address: string): Promise<string | null> {
    return Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const range = ws.getRange(address);
      range.load(['rowCount', 'columnCount']);
      const image = range.getImage();
      await context.sync();

      const numCols = range.columnCount;
      const numRows = range.rowCount;

      // Load column widths and row heights for header compositing
      const cols: Excel.Range[] = [];
      const rows: Excel.Range[] = [];
      for (let i = 0; i < numCols; i++) {
        const col = range.getColumn(i);
        col.format.load('columnWidth');
        cols.push(col);
      }
      for (let i = 0; i < numRows; i++) {
        const row = range.getRow(i);
        row.format.load('rowHeight');
        rows.push(row);
      }
      await context.sync();

      const colWidths = cols.map(c => c.format.columnWidth);
      const rowHeights = rows.map(r => r.format.rowHeight);
      const base64 = image.value;

      // Parse A1 notation to get start row/col indices
      const match = address.replace(/[^A-Za-z0-9:]/g, '').match(/^([A-Za-z]+)(\d+)/);
      let startCol = 0;
      let startRow = 0;
      if (match) {
        const letters = match[1].toUpperCase();
        for (let i = 0; i < letters.length; i++) {
          startCol = startCol * 26 + (letters.charCodeAt(i) - 64);
        }
        startCol -= 1;
        startRow = parseInt(match[2], 10) - 1;
      }

      // Composite headers onto the image using canvas
      return new Promise<string | null>((resolve) => {
        const img = new Image();
        img.onload = () => {
          const HEADER_W = 40;
          const HEADER_H = 20;
          const canvas = document.createElement('canvas');
          canvas.width = HEADER_W + img.width;
          canvas.height = HEADER_H + img.height;
          const ctx = canvas.getContext('2d');
          if (!ctx) { resolve(base64); return; }

          const totalColW = colWidths.reduce((a, b) => a + b, 0);
          const totalRowH = rowHeights.reduce((a, b) => a + b, 0);
          const scaleX = totalColW > 0 ? img.width / totalColW : 1;
          const scaleY = totalRowH > 0 ? img.height / totalRowH : 1;

          ctx.fillStyle = '#ffffff';
          ctx.fillRect(0, 0, canvas.width, canvas.height);
          ctx.drawImage(img, HEADER_W, HEADER_H);

          // Column letter headers
          ctx.fillStyle = '#f0f0f0';
          ctx.fillRect(HEADER_W, 0, img.width, HEADER_H);
          ctx.font = 'bold 10px Calibri, Arial, sans-serif';
          ctx.fillStyle = '#333';
          ctx.textAlign = 'center';
          ctx.textBaseline = 'middle';
          let x = HEADER_W;
          for (let i = 0; i < colWidths.length; i++) {
            const w = colWidths[i] * scaleX;
            let letter = '';
            let tmp = startCol + i;
            do {
              letter = String.fromCharCode((tmp % 26) + 65) + letter;
              tmp = Math.floor(tmp / 26) - 1;
            } while (tmp >= 0);
            ctx.fillStyle = '#333';
            ctx.fillText(letter, x + w / 2, HEADER_H / 2);
            ctx.strokeStyle = '#c0c0c0';
            ctx.strokeRect(x, 0, w, HEADER_H);
            x += w;
          }

          // Row number headers
          ctx.fillStyle = '#f0f0f0';
          ctx.fillRect(0, HEADER_H, HEADER_W, img.height);
          let y = HEADER_H;
          for (let i = 0; i < rowHeights.length; i++) {
            const h = rowHeights[i] * scaleY;
            ctx.fillStyle = '#333';
            ctx.fillText(String(startRow + i + 1), HEADER_W / 2, y + h / 2);
            ctx.strokeStyle = '#c0c0c0';
            ctx.strokeRect(0, y, HEADER_W, h);
            y += h;
          }

          // Corner cell
          ctx.fillStyle = '#f0f0f0';
          ctx.fillRect(0, 0, HEADER_W, HEADER_H);
          ctx.strokeStyle = '#c0c0c0';
          ctx.strokeRect(0, 0, HEADER_W, HEADER_H);

          resolve(canvas.toDataURL('image/png').split(',')[1]);
        };
        img.onerror = () => resolve(base64);
        img.src = `data:image/png;base64,${base64}`;
      });
    });
  }

  // ── From office-agents: search_data ──────────────────────────────────────
  /**
   * Search for a text/value pattern across the workbook (or a specific range).
   * Returns up to maxResults matching cells with address + value.
   */
  static async searchData(
    searchTerm: string,
    options: { matchCase?: boolean; useRegex?: boolean; range?: string; maxResults?: number } = {}
  ): Promise<{ address: string; value: any; sheetName: string }[]> {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();

      const results: { address: string; value: any; sheetName: string }[] = [];
      const max = options.maxResults ?? 200;
      const regex = options.useRegex
        ? new RegExp(searchTerm, options.matchCase ? '' : 'i')
        : null;
      const term = options.matchCase ? searchTerm : searchTerm.toLowerCase();

      for (const sheet of sheets.items) {
        if (results.length >= max) break;
        try {
          const usedRange = options.range
            ? sheet.getRange(options.range)
            : sheet.getUsedRange();
          usedRange.load(['values', 'address', 'rowCount', 'columnCount']);
          await context.sync();

          const _baseAddr = usedRange.address.split('!')[0];
          const startMatch = usedRange.address.match(/\$?([A-Z]+)\$?(\d+)/i);
          const _startCol = startMatch ? startMatch[1] : 'A';
          const startRow = startMatch ? parseInt(startMatch[2], 10) : 1;

          for (let r = 0; r < usedRange.rowCount; r++) {
            for (let c = 0; c < usedRange.columnCount; c++) {
              if (results.length >= max) break;
              const cell = usedRange.values[r][c];
              const cellStr = String(cell ?? '');
              const matches = regex
                ? regex.test(cellStr)
                : (options.matchCase ? cellStr : cellStr.toLowerCase()).includes(term);
              if (matches) {
                // Build cell address
                let colLetter = '';
                let colIdx = c;
                do {
                  colLetter = String.fromCharCode((colIdx % 26) + 65) + colLetter;
                  colIdx = Math.floor(colIdx / 26) - 1;
                } while (colIdx >= 0);
                results.push({
                  address: `${sheet.name}!${colLetter}${startRow + r}`,
                  value: cell,
                  sheetName: sheet.name,
                });
              }
            }
          }
        } catch {
          // Empty sheet — skip
        }
      }
      return results;
    });
  }

  // ── From office-agents: get_all_objects ───────────────────────────────────
  /** List all charts and pivot tables in the workbook. */
  static async getAllObjects(): Promise<{ type: string; name: string; sheetName: string }[]> {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();

      const objects: { type: string; name: string; sheetName: string }[] = [];

      for (const sheet of sheets.items) {
        const charts = sheet.charts;
        const pivots = sheet.pivotTables;
        const tables = sheet.tables;
        charts.load('items/name');
        pivots.load('items/name');
        tables.load('items/name');
        await context.sync();
        charts.items.forEach(c => objects.push({ type: 'chart', name: c.name, sheetName: sheet.name }));
        pivots.items.forEach(p => objects.push({ type: 'pivotTable', name: p.name, sheetName: sheet.name }));
        tables.items.forEach(t => objects.push({ type: 'table', name: t.name, sheetName: sheet.name }));
      }
      return objects;
    });
  }

  // ── From office-agents: get_range_as_csv (token-efficient read) ──────────
  /** Read a range as CSV text — much more token-efficient than JSON for large data. */
  static async getRangeAsCsv(address: string, maxRows: number = 500): Promise<string> {
    return Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const range = ws.getRange(address);
      range.load(['values', 'rowCount']);
      await context.sync();

      const rows = range.values.slice(0, maxRows);
      return rows
        .map(row =>
          row.map(cell => {
            const s = String(cell ?? '');
            return s.includes(',') || s.includes('"') || s.includes('\n')
              ? `"${s.replace(/"/g, '""')}"`
              : s;
          }).join(',')
        )
        .join('\n');
    });
  }

  // ── Freeze panes ──────────────────────────────────────────────────────────
  /** Freeze rows above and/or columns left of the given cell (e.g. "B2" freezes row 1 + col A). */
  static async freezePanes(frozenAtCell: string): Promise<void> {
    return Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      ws.freezePanes.freezeAt(ws.getRange(frozenAtCell));
      await context.sync();
    });
  }

  /** Unfreeze all panes on the active sheet. */
  static async unfreezePanes(): Promise<void> {
    return Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      ws.freezePanes.unfreeze();
      await context.sync();
    });
  }

  // ── Auto-fit columns/rows ─────────────────────────────────────────────────
  /** Auto-fit column widths for a range (makes data readable after writing). */
  static async autoFitColumns(address?: string): Promise<void> {
    return Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const range = address ? ws.getRange(address) : ws.getUsedRange();
      range.format.autofitColumns();
      await context.sync();
    });
  }

  /** Auto-fit row heights for a range. */
  static async autoFitRows(address?: string): Promise<void> {
    return Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const range = address ? ws.getRange(address) : ws.getUsedRange();
      range.format.autofitRows();
      await context.sync();
    });
  }

  // ── Named range management ─────────────────────────────────────────────────
  /** Create or update a named range. */
  static async upsertNamedRange(name: string, address: string): Promise<void> {
    return Excel.run(async (context) => {
      const names = context.workbook.names;
      names.load('items/name');
      await context.sync();
      const existing = names.items.find(n => n.name === name);
      if (existing) {
        existing.delete();
        await context.sync();
      }
      names.add(name, address);
      await context.sync();
    });
  }

  /** Delete a named range. */
  static async deleteNamedRange(name: string): Promise<void> {
    return Excel.run(async (context) => {
      const item = context.workbook.names.getItemOrNullObject(name);
      await context.sync();
      if (!item.isNullObject) {
        item.delete();
        await context.sync();
      }
    });
  }
}

