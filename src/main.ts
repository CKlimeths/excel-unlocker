#!/usr/bin/env node

import { Command } from 'commander';
import * as path from 'path';
import * as fs from 'fs';
import { processExcelFile } from './utils/xlsx-handler.js';
import { ExcelProcessingResult } from './types/index.js';

const program = new Command();

program
  .name('excel-unlocker')
  .description('xlsx工作表保护移除工具 - 支持工作表保护移除和添加')
  .version('1.0.0');

program
  .command('remove')
  .description('移除xlsx文件的工作表保护')
  .requiredOption('-i, --input <file>', '输入xlsx文件路径')
  .option('-o, --output <file>', '输出文件路径（默认：原文件名+"_new.xlsx"）')
  .action(async (options) => {
    try {
      const result = await processExcelFile(options.input, options.output, 'remove');
      console.log(`✅ 处理成功: ${result.inputFile}`);
      console.log(`   输出文件: ${result.outputFile}`);
      console.log(`   操作类型: ${result.operation}`);
      if (result.protected) {
        console.log(`   原文件状态: 受保护`);
      } else {
        console.log(`   原文件状态: 未受保护`);
      }
    } catch (error) {
      console.error(`❌ 处理失败: ${(error as Error).message}`);
      process.exit(1);
    }
  });

program
  .command('add')
  .description('为xlsx文件添加工作表保护')
  .requiredOption('-i, --input <file>', '输入xlsx文件路径')
  .option('-o, --output <file>', '输出文件路径（默认：原文件名+"_new.xlsx"）')
  .option('-p, --password <password>', '保护密码（如未提供，将提示输入）')
  .action(async (options) => {
    try {
      let password = options.password;

      // 如果未提供密码，提示用户输入
      if (!password) {
        const readline = require('readline').createInterface({
          input: process.stdin,
          output: process.stdout
        });

        password = await new Promise((resolve) => {
          readline.question('请输入工作表保护密码: ', (input: string) => {
            readline.close();
            resolve(input);
          });
        });

        if (!password) {
          console.error('❌ 必须提供密码');
          process.exit(1);
        }
      }

      const result = await processExcelFile(options.input, options.output, 'add', password);
      console.log(`✅ 处理成功: ${result.inputFile}`);
      console.log(`   输出文件: ${result.outputFile}`);
      console.log(`   操作类型: ${result.operation}`);
      console.log(`   保护状态: 已添加保护`);
    } catch (error) {
      console.error(`❌ 处理失败: ${(error as Error).message}`);
      process.exit(1);
    }
  });

// 处理拖放操作：如果第一个参数是文件路径，则自动检测并处理
function handleDragAndDrop() {
  const args = process.argv.slice(2);

  // 如果没有参数，显示帮助
  if (args.length === 0) {
    program.help();
    return;
  }

  // 检查第一个参数是否为文件路径
  const possibleFile = args[0];
  if (possibleFile && !possibleFile.startsWith('-') && possibleFile.endsWith('.xlsx')) {
    const filePath = path.resolve(possibleFile);

    if (fs.existsSync(filePath)) {
      console.log(`📁 检测到文件拖放: ${filePath}`);
      console.log('⏳ 正在分析文件状态...');

      // 这里可以自动检测文件是否受保护，然后决定执行remove还是add
      // 目前简化处理：假设用户想要移除保护
      processExcelFile(filePath, undefined, 'detect')
        .then(result => {
          console.log(`✅ 自动处理完成: ${result.inputFile}`);
          console.log(`   输出文件: ${result.outputFile}`);
          console.log(`   执行操作: ${result.operation}`);
          if (result.protected) {
            console.log(`   原文件状态: 受保护 → 已移除保护`);
          } else {
            console.log(`   原文件状态: 未受保护 → 未执行操作（使用 add 命令添加保护）`);
          }
        })
        .catch(error => {
          console.error(`❌ 自动处理失败: ${error.message}`);
          console.log('\n💡 使用以下命令手动处理:');
          console.log('   excel-unlocker remove --input <file>  # 移除保护');
          console.log('   excel-unlocker add --input <file>     # 添加保护');
          process.exit(1);
        });
      return;
    }
  }

  // 如果不是文件拖放，正常解析命令行参数
  program.parse(args);
}

// 主函数
async function main() {
  // 检查是否是文件拖放操作
  if (process.argv.length === 3 && process.argv[2]!.endsWith('.xlsx')) {
    handleDragAndDrop();
  } else {
    program.parse(process.argv);
  }
}

// 错误处理
process.on('uncaughtException', (error) => {
  console.error(`❌ 未处理的错误: ${error.message}`);
  process.exit(1);
});

process.on('unhandledRejection', (error: Error) => {
  console.error(`❌ 未处理的Promise拒绝: ${error.message}`);
  process.exit(1);
});

// 启动程序
if (require.main === module) {
  main().catch(error => {
    console.error(`❌ 程序启动失败: ${error.message}`);
    process.exit(1);
  });
}

export { program };