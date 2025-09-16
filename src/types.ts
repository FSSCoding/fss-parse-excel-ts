export interface ExcelConfig {
  readAllSheets?: boolean;
  includeMetadata?: boolean;
  outputFormat?: 'json' | 'csv' | 'tsv' | 'yaml' | 'markdown';
  safetyChecks?: boolean;
  preserveFormulas?: boolean;
  includeFormatting?: boolean;
}

export interface CellData {
  value: any;
  formula?: string;
  type?: 'string' | 'number' | 'date' | 'boolean' | 'error';
  format?: string;
}

export interface SheetData {
  name: string;
  data: CellData[][];
  range: string;
  rowCount: number;
  columnCount: number;
  metadata?: Record<string, any>;
}

export interface WorkbookData {
  sheets: Record<string, SheetData>;
  metadata?: WorkbookMetadata;
  activeSheet?: string;
}

export interface WorkbookMetadata {
  title?: string;
  author?: string;
  subject?: string;
  creator?: string;
  created?: Date;
  modified?: Date;
  application?: string;
  version?: string;
  sheetNames?: string[];
}

export interface ProcessingResult {
  success: boolean;
  data?: WorkbookData;
  metadata?: WorkbookMetadata;
  warnings?: string[];
  errors?: string[];
  processingTime?: number;
}

export interface SafetyResult {
  isSafe: boolean;
  issues: string[];
  hash: string;
  fileSize: number;
}

export interface ConversionOptions {
  inputPath: string;
  outputPath?: string;
  format: 'json' | 'csv' | 'tsv' | 'yaml' | 'markdown';
  config?: ExcelConfig;
  sheetName?: string;
}

export interface CsvOptions {
  delimiter?: string;
  quote?: string;
  escape?: string;
  header?: boolean;
  skipEmptyLines?: boolean;
}