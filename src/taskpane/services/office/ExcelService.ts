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
        const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        conditionalFormat.colorScale.criteria = {
          minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
          midpoint: { type: Excel.ConditionalFormatColorCriterionType.percentile, formula: "50", color: "#FFEB84" },
          maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" }
        };
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
      
      const toProper = (str: string) => str.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
      
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

  static async getContextForAI(maxCells: number = 1000): Promise<string> {
    try {
      const data = await this.getSelectedRange();
      if (data.rowCount * data.columnCount === 1 && !data.values[0][0]) {
        const used = await this.getUsedRange();
        return `Active Sheet: ${used.sheetName}\nUsed Range: ${used.address}\nData:\n${JSON.stringify(used.values.slice(0, Math.min(used.rowCount, 50)))}`;
      }
      return `Active Sheet: ${data.sheetName}\nSelected Range: ${data.address}\nData:\n${JSON.stringify(data.values)}`;
    } catch (e) {
      return `Unable to read Excel context: ${(e as Error).message}`;
    }
  }
}
