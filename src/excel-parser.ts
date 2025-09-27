import * as fs from 'fs';
import * as XLSX from 'xlsx';
import { parse as csvParse } from 'csv-parse';
import * as yaml from 'js-yaml';
import { ExcelConfig, WorkbookData, SheetData, CellData, ProcessingResult } from './types';
import { SafetyManager } from './safety-manager';

export class ExcelParser {
  private config: ExcelConfig;
  private safetyManager: SafetyManager;

  constructor(config: ExcelConfig = {}) {
    this.config = {
      readAllSheets: true,
      includeMetadata: true,
      outputFormat: 'json',
      safetyChecks: true,
      preserveFormulas: false,
      includeFormatting: false,
      ...config
    };
    
    this.safetyManager = new SafetyManager();
  }

  async parseFile(filePath: string): Promise<ProcessingResult> {
    const startTime = Date.now();
    const result: ProcessingResult = {
      success: false,
      warnings: [],
      errors: []
    };

    try {
      // Safety validation
      if (this.config.safetyChecks) {
        const safetyResult = await this.safetyManager.validateFile(filePath);
        if (!safetyResult.isSafe) {
          result.errors = safetyResult.issues;
          return result;
        }
      }

      const extension = filePath.toLowerCase().split('.').pop();
      
      let workbookData: WorkbookData;
      
      switch (extension) {
        case 'xlsx':
        case 'xls':
        case 'xlsm':
          workbookData = await this.parseExcelFile(filePath);
          break;
        case 'csv':
          workbookData = await this.parseCsvFile(filePath);
          break;
        case 'tsv':
          workbookData = await this.parseTsvFile(filePath);
          break;
        default:
          throw new Error(`Unsupported file format: ${extension}`);
      }

      result.data = workbookData;
      if (workbookData.metadata) {
        result.metadata = workbookData.metadata;
      }
      result.success = true;
      result.processingTime = Date.now() - startTime;

    } catch (error) {
      result.errors?.push(`Parsing failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }

    return result;
  }

  private async parseExcelFile(filePath: string): Promise<WorkbookData> {
    const options: any = {};
    if (this.config.preserveFormulas !== undefined) options.cellFormula = this.config.preserveFormulas;
    if (this.config.includeFormatting !== undefined) options.cellStyles = this.config.includeFormatting;
    // Note: bookProps breaks SheetNames, so we'll extract metadata differently
    
    const workbook = XLSX.readFile(filePath, options);

    const workbookData: WorkbookData = {
      sheets: {}
    };

    // Extract metadata
    if (this.config.includeMetadata && workbook.Props) {
      const metadata: any = {};
      if (workbook.Props.Title) metadata.title = workbook.Props.Title;
      if (workbook.Props.Author) metadata.author = workbook.Props.Author;
      if (workbook.Props.Subject) metadata.subject = workbook.Props.Subject;
      if (workbook.Props.CreatedDate) metadata.created = workbook.Props.CreatedDate;
      if (workbook.Props.ModifiedDate) metadata.modified = workbook.Props.ModifiedDate;
      if (workbook.Props.Application) metadata.application = workbook.Props.Application;
      metadata.sheetNames = workbook.SheetNames;
      workbookData.metadata = metadata;
    }

    // Process sheets
    if (!workbook.SheetNames || !Array.isArray(workbook.SheetNames) || workbook.SheetNames.length === 0) {
      throw new Error('No sheets found in workbook or invalid sheet names');
    }
    
    const sheetsToProcess = this.config.readAllSheets ? workbook.SheetNames : [workbook.SheetNames[0]];
    
    for (const sheetName of sheetsToProcess) {
      const worksheet = workbook.Sheets[sheetName];
      if (worksheet) {
        workbookData.sheets[sheetName] = this.processWorksheet(worksheet, sheetName);
      }
    }

    return workbookData;
  }

  private processWorksheet(worksheet: XLSX.WorkSheet, sheetName: string): SheetData {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
    const data: CellData[][] = [];

    for (let row = range.s.r; row <= range.e.r; row++) {
      const rowData: CellData[] = [];
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[cellAddress];
        
        if (cell) {
          const cellData: CellData = {
            value: cell.v,
            type: this.mapCellType(cell.t),
          };
          
          if (this.config.preserveFormulas && cell.f) {
            cellData.formula = cell.f;
          }
          
          if (this.config.includeFormatting && cell.z) {
            cellData.format = cell.z;
          }
          
          rowData[col - range.s.c] = cellData;
        } else {
          rowData[col - range.s.c] = { value: null, type: 'string' };
        }
      }
      data[row - range.s.r] = rowData;
    }

    return {
      name: sheetName,
      data,
      range: worksheet['!ref'] || 'A1:A1',
      rowCount: range.e.r - range.s.r + 1,
      columnCount: range.e.c - range.s.c + 1
    };
  }

  private mapCellType(xlsxType: string): 'string' | 'number' | 'date' | 'boolean' | 'error' {
    switch (xlsxType) {
      case 'n': return 'number';
      case 'd': return 'date';
      case 'b': return 'boolean';
      case 'e': return 'error';
      case 's':
      default:
        return 'string';
    }
  }

  private async parseCsvFile(filePath: string, delimiter: string = ','): Promise<WorkbookData> {
    return new Promise((resolve, reject) => {
      const data: CellData[][] = [];
      const fileStream = fs.createReadStream(filePath);
      
      fileStream
        .pipe(csvParse({ 
          delimiter,
          skip_empty_lines: true,
          trim: true
        }))
        .on('data', (row: string[]) => {
          const rowData: CellData[] = row.map(cell => ({
            value: this.parseValue(cell),
            type: this.inferType(cell)
          }));
          data.push(rowData);
        })
        .on('end', () => {
          const sheetData: SheetData = {
            name: 'Sheet1',
            data,
            range: `A1:${XLSX.utils.encode_cell({ r: data.length - 1, c: Math.max(0, (data[0]?.length || 1) - 1) })}`,
            rowCount: data.length,
            columnCount: data[0]?.length || 0
          };

          resolve({
            sheets: { 'Sheet1': sheetData }
          });
        })
        .on('error', reject);
    });
  }

  private async parseTsvFile(filePath: string): Promise<WorkbookData> {
    return this.parseCsvFile(filePath, '\t');
  }

  private parseValue(value: string): any {
    // Try to parse as number
    const numValue = Number(value);
    if (!isNaN(numValue) && isFinite(numValue)) {
      return numValue;
    }
    
    // Try to parse as boolean
    if (value.toLowerCase() === 'true') return true;
    if (value.toLowerCase() === 'false') return false;
    
    // Try to parse as date
    const dateValue = new Date(value);
    if (!isNaN(dateValue.getTime()) && value.match(/\d{4}-\d{2}-\d{2}/)) {
      return dateValue;
    }
    
    return value;
  }

  private inferType(value: string): 'string' | 'number' | 'date' | 'boolean' {
    const numValue = Number(value);
    if (!isNaN(numValue) && isFinite(numValue)) {
      return 'number';
    }
    
    if (value.toLowerCase() === 'true' || value.toLowerCase() === 'false') {
      return 'boolean';
    }
    
    const dateValue = new Date(value);
    if (!isNaN(dateValue.getTime()) && value.match(/\d{4}-\d{2}-\d{2}/)) {
      return 'date';
    }
    
    return 'string';
  }

  async convertToFormat(data: WorkbookData, format: string, options?: any): Promise<string> {
    switch (format.toLowerCase()) {
      case 'json':
        return JSON.stringify(data, null, 2);
      case 'csv':
        return this.convertToCSV(data, options?.sheetName);
      case 'tsv':
        return this.convertToCSV(data, options?.sheetName, '\t');
      case 'yaml':
        return yaml.dump(data);
      case 'markdown':
        return this.convertToMarkdown(data, options?.sheetName);
      default:
        throw new Error(`Unsupported output format: ${format}`);
    }
  }

  private convertToCSV(data: WorkbookData, sheetName?: string, delimiter: string = ','): string {
    const sheet = sheetName ? data.sheets[sheetName] : Object.values(data.sheets)[0];
    if (!sheet) {
      throw new Error('No sheet data available');
    }

    const rows = sheet.data.map(row => 
      row.map(cell => String(cell.value || ''))
    );

    return rows.map(row => 
      row.map(cell => `"${cell.replace(/"/g, '""')}"`).join(delimiter)
    ).join('\n');
  }

  private convertToMarkdown(data: WorkbookData, sheetName?: string): string {
    const timestamp = new Date().toISOString().slice(0, 19).replace('T', ' ');
    const sheets = sheetName ? [sheetName] : Object.keys(data.sheets);
    
    // Extract source filename (use sheet name as fallback)
    const sourceFilename = sheetName || 'spreadsheet.xlsx';
    
    let markdown = `# Excel Analysis: ${sourceFilename}\n`;
    markdown += `*Generated on: ${timestamp}*\n\n`;
    
    // Metadata section
    markdown += `## Metadata\n`;
    if (data.metadata) {
      // Calculate totals
      const totalSheets = Object.keys(data.sheets).length;
      const totalRows = Object.values(data.sheets).reduce((sum, sheet) => sum + sheet.rowCount, 0);
      const maxColumns = Math.max(...Object.values(data.sheets).map(sheet => sheet.columnCount));
      
      markdown += `- **File Size:** Unknown\n`; // File size not available in TypeScript version
      markdown += `- **Sheets:** ${totalSheets} (${Object.keys(data.sheets).slice(0, 3).join(', ')}${totalSheets > 3 ? ', ...' : ''})\n`;
      markdown += `- **Total Rows:** ${totalRows.toLocaleString()}\n`;
      markdown += `- **Total Columns:** ${maxColumns}\n`;
      
      if (data.metadata.author) {
        markdown += `- **Author:** ${data.metadata.author}\n`;
      }
      if (data.metadata.created) {
        markdown += `- **Created:** ${new Date(data.metadata.created).toLocaleDateString()}\n`;
      }
      if (data.metadata.modified) {
        markdown += `- **Modified:** ${new Date(data.metadata.modified).toLocaleDateString()}\n`;
      }
      markdown += `- **Format:** XLSX\n`;
    }
    markdown += `\n`;
    
    // Content section
    markdown += `## Content\n\n`;
    
    for (const currentSheetName of sheets) {
      const sheet = data.sheets[currentSheetName];
      if (!sheet) continue;
      
      // Add sheet heading if multiple sheets
      if (sheets.length > 1) {
        markdown += `### Sheet: ${currentSheetName}\n\n`;
      }
      
      if (sheet.data.length === 0) {
        markdown += '*No data in this sheet*\n\n';
        continue;
      }
      
      // Show row count for large sheets
      const maxRows = 20;
      if (sheet.data.length > maxRows) {
        markdown += `*Showing first ${maxRows} rows of ${sheet.data.length.toLocaleString()} total rows*\n\n`;
      }
      
      // Prepare data (limit to first 20 rows for readability)
      const displayData = sheet.data.slice(0, maxRows);
      
      if (displayData.length > 0) {
        // Create table header (use actual column count)
        const columnCount = Math.max(...displayData.map(row => row.length));
        const headers = Array.from({length: columnCount}, (_, i) => `Column ${i + 1}`);
        
        markdown += `| ${headers.join(' | ')} |\n`;
        markdown += `| ${headers.map(() => '---').join(' | ')} |\n`;
        
        // Add data rows with proper escaping
        displayData.forEach(row => {
          const values = Array.from({length: columnCount}, (_, i) => {
            const cell = row[i];
            if (!cell || cell.value === null || cell.value === undefined) {
              return '';
            }
            // Truncate long values and escape pipe characters
            const value = String(cell.value);
            const truncated = value.length > 50 ? value.slice(0, 50) + '...' : value;
            return truncated.replace(/\|/g, '\\|').replace(/\n/g, ' ');
          });
          markdown += `| ${values.join(' | ')} |\n`;
        });
      }
      
      markdown += `\n`;
    }
    
    // Standardized footer
    markdown += `---\n`;
    markdown += `*Generated by FSS Parse Excel v1.0.0*\n`;
    
    return markdown;
  }

  async writeToFile(data: WorkbookData, outputPath: string, format: string): Promise<void> {
    const extension = format.toLowerCase();
    
    switch (extension) {
      case 'xlsx':
        await this.writeExcelFile(data, outputPath);
        break;
      case 'csv':
      case 'tsv':
      case 'json':
      case 'yaml':
      case 'markdown':
        const content = await this.convertToFormat(data, extension);
        fs.writeFileSync(outputPath, content, 'utf8');
        break;
      default:
        throw new Error(`Unsupported output format: ${extension}`);
    }
  }

  private async writeExcelFile(data: WorkbookData, outputPath: string): Promise<void> {
    const workbook = XLSX.utils.book_new();
    
    for (const [sheetName, sheetData] of Object.entries(data.sheets)) {
      const worksheet = XLSX.utils.aoa_to_sheet(
        sheetData.data.map(row => row.map(cell => cell.value))
      );
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    }

    XLSX.writeFile(workbook, outputPath);
  }

  updateConfig(newConfig: Partial<ExcelConfig>): void {
    this.config = { ...this.config, ...newConfig };
  }

  getConfig(): ExcelConfig {
    return { ...this.config };
  }
}