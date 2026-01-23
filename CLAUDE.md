# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a command-line tool for removing and adding worksheet protection to Excel xlsx files. The tool focuses on **worksheet protection** (restricting editing of sheet contents) rather than workbook encryption (password to open file).

**Key Technologies:**
- TypeScript with Node.js runtime (≥18.0.0)
- ExcelJS library for Excel file manipulation
- Commander for CLI argument parsing
- astra-cli for packaging Windows executables
- UPX for executable compression

## Development Commands

### Core Development
```bash
npm install           # Install dependencies
npm run build        # Compile TypeScript to dist/
npm run watch        # Auto-recompile on changes
npm start            # Run compiled program from dist/main.js
npm run dev          # Run with ts-node for development
```

### Testing & Quality
```bash
npm test             # Run tests (currently placeholder)
# Note: Test infrastructure needs to be implemented
```

### Building & Packaging
```bash
npm run package      # Create Windows executable with astra-cli
npm run compress     # Compress executable with UPX
npm run full-build   # Full pipeline: build + package + compress
```

### GitHub Actions
- Automatically builds on tag pushes (`v*`) and pull requests to `main`
- Runs tests, builds TypeScript, packages with astra-cli, compresses with UPX
- Creates GitHub Releases for version tags
- Workflow file: `.github/workflows/build.yml`

## Architecture

### Code Structure
```
src/
├── main.ts                    # CLI entry point - argument parsing, user interaction
├── utils/xlsx-handler.ts      # Core Excel processing logic
└── types/index.ts            # TypeScript type definitions
dist/                         # Compiled JavaScript output
tests/                        # Test files (to be implemented)
```

### Key Components

1. **CLI Interface** (`src/main.ts`):
   - Uses `commander` for command parsing
   - Two main commands: `remove` (remove protection) and `add` (add protection)
   - Supports both command-line arguments and file drag-and-drop
   - Interactive password input for `add` command

2. **Excel Processing** (`src/utils/xlsx-handler.ts`):
   - File validation (existence, format, permissions)
   - Protection detection via ExcelJS `worksheet.protected` property
   - Protection removal via `worksheet.unprotect()`
   - Protection addition via `worksheet.protect(password, options)`
   - Batch processing of all worksheets in workbook

3. **Type System** (`src/types/index.ts`):
   - `OperationType`: 'remove' | 'add' | 'detect' | 'none'
   - `ExcelProcessingResult`: Standardized result object
   - `WorksheetProtectionOptions`: ExcelJS protection configuration

### Data Flow
1. User invokes CLI with file path or drags file onto executable
2. `main.ts` parses arguments and calls `processExcelFile()`
3. `xlsx-handler.ts` validates file, loads with ExcelJS
4. Based on operation type, either removes or adds worksheet protection
5. Saves new file with `_new.xlsx` suffix (or custom output path)
6. Returns structured result with operation details

## Technical Notes

### ExcelJS Integration
- Worksheet protection status: `(worksheet as any).protected === true`
- Remove protection: `worksheet.unprotect()`
- Add protection: `worksheet.protect(password?, options)`
- Options control which operations are allowed (formatting, inserting rows, etc.)
- Empty password string `''` creates protection without password

### TypeScript Configuration
- Target: ES2022 for Node.js compatibility
- Module: nodenext for ESM/CommonJS hybrid
- Strict type checking enabled
- Source maps and declaration files generated
- Output to `dist/` from `src/` root directory

### Packaging & Distribution
- Uses `astra-cli` to bundle Node.js application as Windows executable
- UPX compression reduces executable size but may trigger antivirus warnings
- GitHub Actions automates build pipeline for releases
- Executable designed for Windows but Node.js source is cross-platform

## Common Tasks

### Adding New CLI Options
1. Update command definition in `src/main.ts`
2. Add appropriate validation in `src/utils/xlsx-handler.ts`
3. Update type definitions in `src/types/index.ts` if needed

### Modifying Protection Options
- Edit `options` object in `addWorksheetProtection()` function
- Refer to ExcelJS documentation for available protection settings

### Testing Excel Operations
- Currently no test infrastructure - need to implement test framework
- Consider using test xlsx files with known protection states

### Debugging ExcelJS Issues
- Check ExcelJS version compatibility for protection features
- Some properties may need type assertions: `(worksheet as any).protected`
- ExcelJS methods may throw errors for malformed xlsx files

## Important Constraints

1. **File Format**: Only supports `.xlsx`, not `.xls` or other Excel formats
2. **Protection Type**: Handles worksheet protection only, not workbook encryption
3. **Platform Focus**: Packaging targets Windows, but Node.js source works anywhere
4. **Password Security**: Avoid passing passwords in command-line history

## File References

- **Build Configuration**: `package.json` scripts, `tsconfig.json`
- **CI/CD Pipeline**: `.github/workflows/build.yml`
- **User Documentation**: `README.md` (comprehensive usage instructions)
- **Project Guidelines**: This file (`CLAUDE.md`)