# xlsx工作表保护移除软件

一个用于处理Excel xlsx文件工作表保护（worksheet protection）的命令行工具，支持保护移除和添加功能。

## 功能特性

- **工作表保护移除**：移除xlsx文件中的工作表保护密码
- **工作表保护添加**：为xlsx文件添加工作表保护密码
- **文件格式**：仅支持xlsx格式文件
- **交互方式**：支持命令行参数和文件拖放操作
- **输出文件**：保存为原文件名+"_new.xlsx"

### 注意：工作簿加密 vs 工作表保护
本工具处理的是**工作表保护**（worksheet protection），不是工作簿加密（workbook encryption）：
- **工作表保护**：限制编辑工作表内容，但可以打开文件查看
- **工作簿加密**：需要密码才能打开整个文件
本工具专注于工作表保护操作

## 安装方法

### 从源码构建
```bash
# 克隆仓库
git clone <repository-url>
cd excel-unlocker

# 安装依赖
bun install

# 构建项目
bun run build
```

构建完成后，可执行文件位于 `dist/excel-unlocker.exe`。

## 使用方法

### 方式一：命令行参数
```bash
# 移除工作表保护
excel-unlocker remove --input 文件.xlsx

# 添加工作表保护（交互式输入密码）
excel-unlocker add --input 文件.xlsx

# 添加工作表保护（指定密码）
excel-unlocker add --input 文件.xlsx --password "你的密码"

# 指定输出文件
excel-unlocker remove --input 文件.xlsx --output 文件_无保护.xlsx

# 显示帮助信息
excel-unlocker --help
```

### 方式二：文件拖放操作
将xlsx文件拖放到 `excel-unlocker.exe` 图标上，程序会自动检测文件状态：
- 如果文件有工作表保护，则移除保护
- 如果文件没有保护，则提示输入密码并添加保护
输出文件保存为原文件名+"_new.xlsx"。

### 批量处理
支持使用通配符处理多个文件：
```bash
excel-unlocker remove --input "*.xlsx"
```

## 技术实现

### 技术栈
- **开发语言**：TypeScript
- **运行环境**：Bun
- **Excel处理**：exceljs库
- **命令行解析**：commander库
- **打包工具**：Bun
- **目标平台**：Windows(x64)

### 项目结构
```
excel-unlocker/
├── src/                    # 源代码
│   ├── main.ts             # 主程序入口
│   ├── utils/
│   │   └── xlsx-handler.ts # Excel处理逻辑
│   └── types/
│       └── index.ts        # 类型定义
├── tests/                  # 测试文件
├── dist/                   # 构建输出
├── package.json            # 项目配置
├── tsconfig.json           # TypeScript配置
└── README.md               # 项目说明
```

## 开发指南

### 环境准备
安装 Bun `winget install --id Oven-sh.Bun`

### 开发命令
```bash
# 安装依赖
bun install

# 构建项目
bun run build

# 运行程序
bun start
```

## 构建与部署

### GitHub Actions 自动化构建
项目配置了 GitHub Actions 工作流，自动打包为 Windows 可执行文件

### 手动构建
```bash
# 安装依赖
bun ci

# 构建二进制包
bun run build
```
