#!/usr/bin/env node

import { Command } from 'commander';
import chalk from 'chalk';
import ora from 'ora';
import * as fs from 'fs';
import * as path from 'path';
import { ExcelParser } from './excel-parser';
import { ExcelConfig } from './types';

const program = new Command();

program
  .name('fss-parse-excel-ts')
  .description('Professional TypeScript Excel/spreadsheet parsing and manipulation toolkit')
  .version('1.0.0');

program
  .command('parse')
  .description('Parse Excel/CSV file and extract data')
  .argument('<input>', 'Input file path (XLSX, XLS, CSV, TSV)')
  .option('-o, --output <path>', 'Output file path')
  .option('-f, --format <format>', 'Output format (json|csv|tsv|yaml|markdown)', 'json')
  .option('-s, --sheet <name>', 'Specific sheet name to process')
  .option('--all-sheets', 'Process all sheets')
  .option('--no-metadata', 'Skip metadata extraction')
  .option('--formulas', 'Preserve formulas')
  .option('--formatting', 'Include cell formatting')
  .action(async (input: string, options: any) => {
    const spinner = ora('Parsing spreadsheet...').start();

    try {
      const config: ExcelConfig = {
        readAllSheets: options.allSheets,
        includeMetadata: options.metadata !== false,
        preserveFormulas: options.formulas,
        includeFormatting: options.formatting,
        outputFormat: options.format
      };

      const parser = new ExcelParser(config);
      const result = await parser.parseFile(input);

      if (!result.success) {
        spinner.fail('Parsing failed');
        if (result.errors?.length) {
          console.error(chalk.red('Errors:'));
          result.errors.forEach(error => console.error(chalk.red(`  • ${error}`)));
        }
        process.exit(1);
      }

      if (result.data) {
        const output = await parser.convertToFormat(result.data, options.format, {
          sheetName: options.sheet
        });
        
        if (options.output) {
          fs.writeFileSync(options.output, output, 'utf8');
          spinner.succeed(`Data extracted to ${chalk.green(options.output)}`);
        } else {
          spinner.stop();
          console.log(output);
        }
      }

      // Show metadata
      if (result.metadata && options.format !== 'json') {
        console.log(chalk.blue('\nSpreadsheet Information:'));
        if (result.metadata.title) console.log(`  Title: ${result.metadata.title}`);
        if (result.metadata.author) console.log(`  Author: ${result.metadata.author}`);
        if (result.metadata.sheetNames) {
          console.log(`  Sheets: ${result.metadata.sheetNames.join(', ')}`);
        }
      }

      if (result.processingTime) {
        console.log(chalk.gray(`\nProcessed in ${result.processingTime}ms`));
      }

    } catch (error) {
      spinner.fail('Processing failed');
      console.error(chalk.red(error instanceof Error ? error.message : 'Unknown error'));
      process.exit(1);
    }
  });

program
  .command('convert')
  .description('Convert between Excel formats')
  .argument('<input>', 'Input file path')
  .argument('<output>', 'Output file path')
  .option('-s, --sheet <name>', 'Specific sheet to convert')
  .option('--formulas', 'Preserve formulas in output')
  .action(async (input: string, output: string, options: any) => {
    const spinner = ora(`Converting ${path.basename(input)}...`).start();

    try {
      const parser = new ExcelParser({
        preserveFormulas: options.formulas
      });
      
      const result = await parser.parseFile(input);
      
      if (!result.success || !result.data) {
        spinner.fail('Conversion failed');
        if (result.errors?.length) {
          result.errors.forEach(error => console.error(chalk.red(error)));
        }
        process.exit(1);
      }

      const format = path.extname(output).slice(1).toLowerCase();
      await parser.writeToFile(result.data, output, format);
      
      spinner.succeed(`Converted to ${chalk.green(output)}`);

    } catch (error) {
      spinner.fail('Conversion failed');
      console.error(chalk.red(error instanceof Error ? error.message : 'Unknown error'));
      process.exit(1);
    }
  });

program
  .command('info')
  .description('Display Excel file information')
  .argument('<input>', 'Input file path')
  .option('--detailed', 'Show detailed sheet information')
  .action(async (input: string, options: any) => {
    const spinner = ora('Reading file information...').start();

    try {
      const parser = new ExcelParser({ includeMetadata: true });
      const result = await parser.parseFile(input);

      if (!result.success || !result.data) {
        spinner.fail('Failed to read file');
        if (result.errors?.length) {
          result.errors.forEach(error => console.error(chalk.red(error)));
        }
        process.exit(1);
      }

      spinner.stop();

      console.log(chalk.blue.bold(`\n📊 ${path.basename(input)}`));
      console.log(chalk.gray('─'.repeat(50)));

      if (result.metadata) {
        if (result.metadata.title) {
          console.log(`${chalk.yellow('Title:')} ${result.metadata.title}`);
        }
        if (result.metadata.author) {
          console.log(`${chalk.yellow('Author:')} ${result.metadata.author}`);
        }
        if (result.metadata.created) {
          console.log(`${chalk.yellow('Created:')} ${result.metadata.created.toLocaleDateString()}`);
        }
        if (result.metadata.modified) {
          console.log(`${chalk.yellow('Modified:')} ${result.metadata.modified.toLocaleDateString()}`);
        }
        if (result.metadata.application) {
          console.log(`${chalk.yellow('Application:')} ${result.metadata.application}`);
        }
      }

      // Sheet information
      const sheets = Object.entries(result.data.sheets);
      console.log(`${chalk.yellow('Sheets:')} ${sheets.length}`);

      if (options.detailed) {
        console.log(chalk.gray('\nSheet Details:'));
        sheets.forEach(([name, sheet]) => {
          console.log(`  📋 ${chalk.cyan(name)}`);
          console.log(`     Rows: ${sheet.rowCount}, Columns: ${sheet.columnCount}`);
          console.log(`     Range: ${sheet.range}`);
        });
      } else {
        const sheetNames = sheets.map(([name]) => name);
        console.log(`  ${sheetNames.join(', ')}`);
      }

    } catch (error) {
      spinner.fail('Information extraction failed');
      console.error(chalk.red(error instanceof Error ? error.message : 'Unknown error'));
      process.exit(1);
    }
  });

program
  .command('validate')
  .description('Validate Excel file safety and integrity')
  .argument('<input>', 'Input file path')
  .action(async (input: string) => {
    const spinner = ora('Validating file...').start();

    try {
      const parser = new ExcelParser({ safetyChecks: true });
      const result = await parser.parseFile(input);

      spinner.stop();

      if (result.success) {
        console.log(chalk.green('✅ File validation passed'));
        
        if (result.warnings?.length) {
          console.log(chalk.yellow('\n⚠️  Warnings:'));
          result.warnings.forEach(warning => console.log(chalk.yellow(`  • ${warning}`)));
        }
      } else {
        console.log(chalk.red('❌ File validation failed'));
        
        if (result.errors?.length) {
          console.log(chalk.red('\nErrors:'));
          result.errors.forEach(error => console.log(chalk.red(`  • ${error}`)));
        }
        process.exit(1);
      }

    } catch (error) {
      spinner.fail('Validation failed');
      console.error(chalk.red(error instanceof Error ? error.message : 'Unknown error'));
      process.exit(1);
    }
  });

program
  .command('extract-sheets')
  .description('Extract specific sheets as separate files')
  .argument('<input>', 'Input Excel file')
  .option('-o, --output-dir <dir>', 'Output directory', './extracted')
  .option('-f, --format <format>', 'Output format (csv|json|yaml)', 'csv')
  .action(async (input: string, options: any) => {
    const spinner = ora('Extracting sheets...').start();

    try {
      const parser = new ExcelParser({ readAllSheets: true });
      const result = await parser.parseFile(input);

      if (!result.success || !result.data) {
        spinner.fail('Extraction failed');
        if (result.errors?.length) {
          result.errors.forEach(error => console.error(chalk.red(error)));
        }
        process.exit(1);
      }

      if (!fs.existsSync(options.outputDir)) {
        fs.mkdirSync(options.outputDir, { recursive: true });
      }

      const sheets = Object.entries(result.data.sheets);
      let extracted = 0;

      for (const [sheetName, sheetData] of sheets) {
        const fileName = `${sheetName.replace(/[^a-zA-Z0-9]/g, '_')}.${options.format}`;
        const outputPath = path.join(options.outputDir, fileName);
        
        const sheetWorkbook: any = { 
          sheets: { [sheetName]: sheetData }
        };
        if (result.data.metadata) {
          sheetWorkbook.metadata = result.data.metadata;
        }
        
        const content = await parser.convertToFormat(sheetWorkbook, options.format, {
          sheetName
        });
        
        fs.writeFileSync(outputPath, content, 'utf8');
        extracted++;
      }

      spinner.succeed(`Extracted ${extracted} sheets to ${chalk.green(options.outputDir)}`);

    } catch (error) {
      spinner.fail('Extraction failed');
      console.error(chalk.red(error instanceof Error ? error.message : 'Unknown error'));
      process.exit(1);
    }
  });

// Global error handlers
process.on('uncaughtException', (error) => {
  console.error(chalk.red('Uncaught Exception:'), error.message);
  process.exit(1);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error(chalk.red('Unhandled Rejection at:'), promise, 'reason:', reason);
  process.exit(1);
});

program.parse();