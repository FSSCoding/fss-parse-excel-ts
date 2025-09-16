import * as fs from 'fs';
import * as crypto from 'crypto';
import * as path from 'path';
import { SafetyResult } from './types';

export class SafetyManager {
  private maxFileSize: number;
  private allowedExtensions: Set<string>;
  private backupDir: string;

  constructor() {
    this.maxFileSize = 500 * 1024 * 1024; // 500MB
    this.allowedExtensions = new Set(['.xlsx', '.xls', '.xlsm', '.csv', '.tsv', '.ods']);
    this.backupDir = path.join(process.cwd(), '.fss-excel-backups');
  }

  async validateFile(filePath: string): Promise<SafetyResult> {
    const issues: string[] = [];

    try {
      if (!fs.existsSync(filePath)) {
        issues.push('File does not exist');
        return { isSafe: false, issues, hash: '', fileSize: 0 };
      }

      const ext = path.extname(filePath).toLowerCase();
      if (!this.allowedExtensions.has(ext)) {
        issues.push(`Unsupported file extension: ${ext}`);
      }

      const stats = fs.statSync(filePath);
      const fileSize = stats.size;

      if (fileSize > this.maxFileSize) {
        issues.push(`File too large: ${fileSize} bytes (max: ${this.maxFileSize})`);
      }

      const fileBuffer = fs.readFileSync(filePath);
      const hash = crypto.createHash('sha256').update(fileBuffer).digest('hex');

      if (await this.basicMalwareCheck(fileBuffer)) {
        issues.push('Potential security risk detected');
      }

      return {
        isSafe: issues.length === 0,
        issues,
        hash,
        fileSize
      };

    } catch (error) {
      issues.push(`Validation error: ${error instanceof Error ? error.message : 'Unknown error'}`);
      return { isSafe: false, issues, hash: '', fileSize: 0 };
    }
  }

  async createBackup(filePath: string): Promise<string> {
    try {
      if (!fs.existsSync(this.backupDir)) {
        fs.mkdirSync(this.backupDir, { recursive: true });
      }

      const fileName = path.basename(filePath);
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const backupPath = path.join(this.backupDir, `${timestamp}-${fileName}`);

      fs.copyFileSync(filePath, backupPath);
      return backupPath;

    } catch (error) {
      throw new Error(`Backup failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  private async basicMalwareCheck(buffer: Buffer): Promise<boolean> {
    const suspiciousPatterns = [
      'eval(',
      'exec(',
      'system(',
      'shell_exec(',
      'javascript:',
      '<script',
      'vbscript:',
      'ActiveXObject'
    ];

    const content = buffer.toString('utf8', 0, Math.min(buffer.length, 10000));
    
    return suspiciousPatterns.some(pattern => 
      content.toLowerCase().includes(pattern.toLowerCase())
    );
  }

  setMaxFileSize(size: number): void {
    this.maxFileSize = size;
  }

  addAllowedExtension(ext: string): void {
    this.allowedExtensions.add(ext.toLowerCase());
  }

  setBackupDirectory(dir: string): void {
    this.backupDir = dir;
  }
}