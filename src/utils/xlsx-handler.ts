import * as ExcelJS from "exceljs"
import * as path from "path"
import * as fs from "fs"
import { ExcelProcessingResult, OperationType } from "@/types/index.js"

/**
 * 检查文件是否是有效的xlsx文件
 */
async function validateExcelFile(filePath: string): Promise<void> {
    if (!fs.existsSync(filePath)) {
        throw new Error(`文件不存在: ${filePath}`)
    }

    const ext = path.extname(filePath).toLowerCase()
    if (ext !== ".xlsx") {
        throw new Error(`不支持的文件格式: ${ext}，仅支持.xlsx文件`)
    }

    // 检查文件是否可读
    try {
        fs.accessSync(filePath, fs.constants.R_OK)
    } catch {
        throw new Error(`文件不可读或没有读取权限: ${filePath}`)
    }
}

/**
 * 检测工作表是否受保护
 */
async function isWorksheetProtected(
    worksheet: ExcelJS.Worksheet,
): Promise<boolean> {
    try {
        // 检查工作表保护状态
        // exceljs中通过worksheet.protected属性检查保护状态
        return (worksheet as any).protected === true
    } catch (error) {
        console.warn(`检测工作表保护状态时出错: ${(error as Error).message}`)
        return false
    }
}

/**
 * 移除工作表保护
 */
async function removeWorksheetProtection(
    worksheet: ExcelJS.Worksheet,
): Promise<void> {
    try {
        // 移除工作表保护 - exceljs中使用unprotect方法
        worksheet.unprotect()
    } catch (error) {
        throw new Error(`移除工作表保护失败: ${(error as Error).message}`)
    }
}

/**
 * 添加工作表保护
 */
async function addWorksheetProtection(
    worksheet: ExcelJS.Worksheet,
    password?: string,
): Promise<void> {
    try {
        // 设置工作表保护 - exceljs中使用protect方法
        const options = {
            // 保护选项
            selectLockedCells: false,
            selectUnlockedCells: true,
            formatCells: false,
            formatColumns: false,
            formatRows: false,
            insertColumns: false,
            insertRows: false,
            insertHyperlinks: false,
            deleteColumns: false,
            deleteRows: false,
            sort: false,
            autoFilter: false,
            pivotTables: false,
        }

        if (password) {
            worksheet.protect(password, options)
        } else {
            worksheet.protect("", options)
        }
    } catch (error) {
        throw new Error(`添加工作表保护失败: ${(error as Error).message}`)
    }
}

/**
 * 处理Excel文件的主函数
 */
export async function processExcelFile(
    inputFilePath: string,
    outputFilePath?: string,
    operation: OperationType = "detect",
    password?: string,
): Promise<ExcelProcessingResult> {
    // 验证输入文件
    await validateExcelFile(inputFilePath)

    // 生成输出文件路径
    let outputPath = outputFilePath
    if (!outputPath) {
        const dir = path.dirname(inputFilePath)
        const ext = path.extname(inputFilePath)
        const baseName = path.basename(inputFilePath, ext)
        outputPath = path.join(dir, `${baseName}_new${ext}`)
    }

    // 确保输出目录存在
    const outputDir = path.dirname(outputPath)
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true })
    }

    // 检查输出文件是否已存在
    if (fs.existsSync(outputPath) && outputPath !== inputFilePath) {
        throw new Error(`输出文件已存在: ${outputPath}，请指定其他输出路径`)
    }

    // 加载Excel工作簿
    const workbook = new ExcelJS.Workbook()

    try {
        await workbook.xlsx.readFile(inputFilePath)
    } catch (error) {
        throw new Error(`读取Excel文件失败: ${(error as Error).message}`)
    }

    let anyProtected = false
    let operationPerformed: OperationType = "none"
    const processedSheets: string[] = []

    // 处理每个工作表
    workbook.eachSheet((worksheet, sheetId) => {
        processedSheets.push(worksheet.name)
    })

    // 重新遍历以执行操作
    for (const sheetName of processedSheets) {
        const worksheet = workbook.getWorksheet(sheetName)
        if (!worksheet) continue

        const isProtected = await isWorksheetProtected(worksheet)

        if (isProtected) {
            anyProtected = true
        }

        // 根据操作类型执行相应操作
        switch (operation) {
            case "remove":
                if (isProtected) {
                    await removeWorksheetProtection(worksheet)
                    operationPerformed = "remove"
                }
                break

            case "add":
                if (!isProtected) {
                    await addWorksheetProtection(worksheet, password)
                    operationPerformed = "add"
                }
                break

            case "detect":
                // 自动检测：如果有保护则移除，否则不操作
                if (isProtected) {
                    await removeWorksheetProtection(worksheet)
                    operationPerformed = "remove"
                }
                break

            default:
                throw new Error(`不支持的操作类型: ${operation}`)
        }
    }

    // 如果没有工作表被处理
    if (
        operation === "remove" &&
        !anyProtected &&
        operationPerformed === "none"
    ) {
        console.log(`ℹ️  文件没有工作表保护，无需移除: ${inputFilePath}`)
    }

    if (operation === "add" && anyProtected && operationPerformed === "none") {
        console.log(`ℹ️  文件已有工作表保护，无需添加: ${inputFilePath}`)
    }

    // 保存工作簿
    try {
        await workbook.xlsx.writeFile(outputPath)
    } catch (error) {
        throw new Error(`保存Excel文件失败: ${(error as Error).message}`)
    }

    return {
        inputFile: inputFilePath,
        outputFile: outputPath,
        operation: operationPerformed,
        protected: anyProtected,
        sheetsProcessed: processedSheets.length,
        success: true,
        timestamp: new Date().toISOString(),
    }
}

/**
 * 检测文件保护状态（不修改文件）
 */
export async function detectProtection(filePath: string): Promise<{
    protected: boolean
    sheetCount: number
    protectedSheets: string[]
}> {
    await validateExcelFile(filePath)

    const workbook = new ExcelJS.Workbook()

    try {
        await workbook.xlsx.readFile(filePath)
    } catch (error) {
        throw new Error(`读取Excel文件失败: ${(error as Error).message}`)
    }

    const protectedSheets: string[] = []
    let sheetCount = 0

    workbook.eachSheet((worksheet, sheetId) => {
        sheetCount++
        const isProtected = (worksheet as any).protected === true
        if (isProtected) {
            protectedSheets.push(worksheet.name)
        }
    })

    return {
        protected: protectedSheets.length > 0,
        sheetCount,
        protectedSheets,
    }
}
