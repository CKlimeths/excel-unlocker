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
本工具专注于工作表保护操作。

## 安装方法

### 从源码构建
```bash
# 克隆仓库
git clone <repository-url>
cd excel-unlocker

# 安装依赖
npm install

# 构建项目
npm run build

# 打包为Windows可执行文件
npm run full-build
```

构建完成后，可执行文件位于 `dist/excel-unlocker.exe`。

### 直接使用可执行文件
从 Releases 页面下载最新版本的 `excel-unlocker.exe`。

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
- **运行环境**：Node.js (>=18.0.0)
- **Excel处理**：exceljs库
- **命令行解析**：commander库
- **打包工具**：astra-cli
- **压缩工具**：upx
- **目标平台**：Windows

### 项目结构
```
excel-unlocker/
├── src/                    # 源代码
│   ├── main.ts            # 主程序入口
│   ├── utils/
│   │   └── xlsx-handler.ts # Excel处理逻辑
│   └── types/
│       └── index.ts       # 类型定义
├── tests/                 # 测试文件
├── dist/                  # 构建输出
├── package.json          # 项目配置
├── tsconfig.json         # TypeScript配置
└── README.md             # 项目说明
```

## 开发指南

### 环境准备
1. 安装 Node.js (>=18.0.0)
2. 安装 TypeScript: `npm install -g typescript`
3. 安装 astra-cli: `npm install -g astra-cli`
4. 安装 upx (可选，用于压缩可执行文件)

### 开发命令
```bash
# 安装依赖
npm install

# 开发模式（自动重编译）
npm run watch

# 运行测试
npm test

# 构建项目
npm run build

# 运行程序
npm start

# 打包为可执行文件
npm run package

# 完整构建（编译+打包+压缩）
npm run full-build
```

### 测试
项目包含单元测试和集成测试，确保功能稳定性：
```bash
# 运行所有测试
npm test

# 运行特定测试文件
npm test -- tests/xlsx-handler.test.ts
```

## 构建与部署

### GitHub Actions 自动化构建
项目配置了 GitHub Actions 工作流，自动：
1. 运行测试
2. 构建 TypeScript 代码
3. 打包为 Windows 可执行文件
4. 使用 upx 压缩
5. 创建 GitHub Release

### 手动构建
```bash
# 安装依赖
npm install

# 编译TypeScript
npm run build

# 打包为可执行文件
npm run package

# 压缩可执行文件（可选）
npm run compress
```

## 注意事项

### 安全性
- 密码输入：避免在命令行中直接传递密码，建议使用交互式输入
- 文件处理：处理敏感文件时，确保输出文件权限适当
- 临时文件：处理过程中产生的临时文件会被自动清理

### 限制
- 仅支持 xlsx 格式，不支持 xls 格式
- 仅处理工作表保护，不处理工作簿加密
- 需要 Node.js 运行时环境或预构建的可执行文件
- 主要针对 Windows 平台，其他平台可能需要调整

### 错误处理
程序包含完整的错误处理机制：
- 文件不存在或格式错误
- 权限不足
- 密码错误
- 磁盘空间不足

## 许可证
MIT License

## 贡献指南
欢迎提交 Issue 和 Pull Request。在提交代码前，请确保：
1. 代码通过所有测试
2. 遵循项目代码风格
3. 更新相关文档
4. 添加适当的测试用例

## 问题反馈
如遇问题，请：
1. 查看 Issues 页面是否已有类似问题
2. 提供详细的错误信息和复现步骤
3. 附上相关的 Excel 文件（如可能）
