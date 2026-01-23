/**
 * 操作类型定义
 */
export type OperationType = 'remove' | 'add' | 'detect' | 'none';

/**
 * Excel处理结果接口
 */
export interface ExcelProcessingResult {
  /** 输入文件路径 */
  inputFile: string;

  /** 输出文件路径 */
  outputFile: string;

  /** 执行的操作类型 */
  operation: OperationType;

  /** 原文件是否有保护 */
  protected: boolean;

  /** 处理的工作表数量 */
  sheetsProcessed: number;

  /** 是否成功 */
  success: boolean;

  /** 处理时间戳 */
  timestamp: string;
}

/**
 * 命令行选项接口
 */
export interface CommandOptions {
  /** 输入文件路径 */
  input: string;

  /** 输出文件路径（可选） */
  output?: string;

  /** 密码（可选） */
  password?: string;

  /** 详细输出 */
  verbose?: boolean;
}

/**
 * 工作表保护选项接口
 */
export interface WorksheetProtectionOptions {
  /** 是否受保护 */
  protected: boolean;

  /** 密码（可选） */
  password?: string;

  /** 允许选择锁定的单元格 */
  selectLockedCells?: boolean;

  /** 允许选择未锁定的单元格 */
  selectUnlockedCells?: boolean;

  /** 允许格式化单元格 */
  formatCells?: boolean;

  /** 允许格式化列 */
  formatColumns?: boolean;

  /** 允许格式化行 */
  formatRows?: boolean;

  /** 允许插入列 */
  insertColumns?: boolean;

  /** 允许插入行 */
  insertRows?: boolean;

  /** 允许插入超链接 */
  insertHyperlinks?: boolean;

  /** 允许删除列 */
  deleteColumns?: boolean;

  /** 允许删除行 */
  deleteRows?: boolean;

  /** 允许排序 */
  sort?: boolean;

  /** 允许自动筛选 */
  autoFilter?: boolean;

  /** 允许数据透视表 */
  pivotTables?: boolean;
}

/**
 * 错误类型定义
 */
export interface ExcelUnlockerError {
  /** 错误代码 */
  code: string;

  /** 错误消息 */
  message: string;

  /** 错误详情 */
  details?: unknown;

  /** 发生时间 */
  timestamp: string;
}

/**
 * 文件信息接口
 */
export interface FileInfo {
  /** 文件路径 */
  path: string;

  /** 文件大小（字节） */
  size: number;

  /** 修改时间 */
  modified: Date;

  /** 是否受保护 */
  protected: boolean;

  /** 工作表数量 */
  sheetCount: number;
}

/**
 * 批量处理结果接口
 */
export interface BatchProcessingResult {
  /** 总文件数 */
  totalFiles: number;

  /** 成功处理数 */
  successful: number;

  /** 失败数 */
  failed: number;

  /** 跳过数 */
  skipped: number;

  /** 处理结果列表 */
  results: Array<{
    file: string;
    success: boolean;
    operation: OperationType;
    outputFile?: string;
    error?: string;
  }>;

  /** 开始时间 */
  startTime: string;

  /** 结束时间 */
  endTime: string;

  /** 总耗时（毫秒） */
  duration: number;
}