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
      if (type.toLowerCase().includes('pie')) chartType = Excel.ChartType.pie;
      else if (type.toLowerCase().includes('line')) chartType = Excel.ChartType.line;
      else if (type.toLowerCase().includes('bar')) chartType = Excel.ChartType.barClustered;

      const chart = worksheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);
      chart.title.text = title;
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
