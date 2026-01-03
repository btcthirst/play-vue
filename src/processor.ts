import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { Constants } from './constants';

export interface ProcessedFile {
    name: string;
    buffer: ArrayBuffer;
}

export async function processNarahuvannya(file: File, sourceYear: string, targetYear: string): Promise<ProcessedFile> {
    const buffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    for (let i = 0; i < Constants.Months.length; i++) {
        const month = Constants.Months[i];
        const sheet = workbook.getWorksheet(month);
        if (!sheet) continue;

        // Data copying/referencing logic
        if (month === Constants.SheetSichen) {
            const prevSheet = workbook.getWorksheet("грудень");
            if (prevSheet) {
                for (let r = Constants.StartRowData; r <= Constants.EndRowNarahuvannya; r++) {
                    // B -> B
                    copyCell(prevSheet, sheet, `B${r}`, `B${r}`);
                    // P -> D
                    copyCell(prevSheet, sheet, `P${r}`, `D${r}`);
                    // Q -> E
                    copyCell(prevSheet, sheet, `Q${r}`, `E${r}`);
                }
            }
        } else {
            const prevMonth = Constants.Months[i - 1];
            for (let r = Constants.StartRowData; r <= Constants.EndRowNarahuvannya; r++) {
                const cell = sheet.getCell(`B${r}`);
                cell.value = { formula: `'${prevMonth}'!B${r}` };
            }
        }

        // Clear columns K, L, M, O
        const colsToClear = ["K", "L", "M", "O"];
        for (let r = Constants.StartRowData; r <= Constants.EndRowNarahuvannya; r++) {
            for (const col of colsToClear) {
                sheet.getCell(`${col}${r}`).value = null;
            }
        }

        // Update year in headers
        for (let r = Constants.HeaderRowStart; r <= Constants.HeaderRowEnd; r++) {
            const row = sheet.getRow(r);
            row.eachCell((cell) => {
                if (typeof cell.value === 'string' && cell.value.includes(sourceYear)) {
                    cell.value = cell.value.replace(sourceYear, targetYear);
                }
            });
        }
    }

    // Update "Лист1"
    const sheet1 = workbook.getWorksheet(Constants.SheetList1);
    if (sheet1) {
        sheet1.eachRow((row) => {
            row.eachCell((cell) => {
                if (typeof cell.value === 'string' && cell.value.includes(sourceYear)) {
                    cell.value = cell.value.replace(sourceYear, targetYear);
                }
            });
        });
    }

    const outBuffer = await workbook.xlsx.writeBuffer();
    return { name: file.name.replace(sourceYear.slice(2), targetYear.slice(2)), buffer: outBuffer };
}

export async function processVidomist(file: File, sourceYear: string, targetYear: string): Promise<ProcessedFile> {
    const buffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    for (let i = 0; i < Constants.Months.length; i++) {
        const month = Constants.Months[i];
        const sheet = workbook.getWorksheet(month);
        if (!sheet) continue;

        // 1. Update S1 in "січень"
        if (month === Constants.SheetSichen) {
            const cell = sheet.getCell(Constants.CellYearLocation);
            if (typeof cell.value === 'string' && cell.value.includes(sourceYear)) {
                cell.value = cell.value.replace(sourceYear, targetYear);
            }
        }

        // 2. Update year in headers/footers (and all cells)
        sheet.eachRow((row) => {
            row.eachCell((cell) => {
                if (typeof cell.value === 'string' && cell.value.includes(sourceYear)) {
                    cell.value = cell.value.replace(sourceYear, targetYear);
                } else if (cell.value && typeof cell.value === 'object' && 'formula' in cell.value) {
                    // Formula handling if needed, though exceljs formula is object
                    // For simple string replacement in formula, we might need to check cell.model.formula
                }
            });
        });

        // 3. Clear data while preserving formulas
        let totalRow = -1;
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber < Constants.StartRowData) return;
            const cellD = row.getCell('D');
            const cellF = row.getCell('F');
            // Check if formula contains SUM (simplified check)
            const isSum = (cell: ExcelJS.Cell) => cell.formula && cell.formula.toUpperCase().includes('SUM');

            if (isSum(cellD) || isSum(cellF)) {
                totalRow = rowNumber;
                return false; // break
            }
        });

        if (totalRow !== -1) {
            const colsToClear = ["A", "B", "C", "D", "G", "H", "I", "J", "K", "L"];
            for (let r = Constants.StartRowData; r < totalRow; r++) {
                for (const col of colsToClear) {
                    const cell = sheet.getCell(`${col}${r}`);
                    if (!cell.formula) {
                        cell.value = null;
                    }
                }
            }
        }

        // 4. Update Balance Labels
        const headerCell = sheet.getCell(Constants.CellBalanceHeader);
        if (month === Constants.SheetSichen) {
            headerCell.value = `Сальдо на 01.01.${targetYear}р`;
        } else {
            const prevMonthName = Constants.MonthNamesFull[i - 1];
            headerCell.value = `Сальдо на ${prevMonthName} ${targetYear}р`;
        }

        const currentMonthName = Constants.MonthNamesFull[i];
        if (totalRow !== -1) {
            for (let r = totalRow; r <= totalRow + 5; r++) {
                const cellE = sheet.getCell(`E${r}`);
                if (typeof cellE.value === 'string' && cellE.value.includes("Сальдо на")) {
                    cellE.value = `Сальдо на ${currentMonthName} ${targetYear}р`;
                    break;
                }
            }
        }
    }

    const outBuffer = await workbook.xlsx.writeBuffer();
    return { name: file.name.replace(sourceYear.slice(2), targetYear.slice(2)), buffer: outBuffer };
}

export async function processBorzhnyky(file: File, sourceYear: string, targetYear: string): Promise<ProcessedFile> {
    const buffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buffer);

    const linkFile = zip.file("xl/externalLinks/_rels/externalLink1.xml.rels");
    if (linkFile) {
        let content = await linkFile.async("string");
        const srcSuffix = sourceYear.slice(2) + ".xlsx";
        const dstSuffix = targetYear.slice(2) + ".xlsx";
        const srcLink = Constants.PrefixNarahuvannya + srcSuffix;
        const dstLink = Constants.PrefixNarahuvannya + dstSuffix;

        content = content.split(srcLink).join(dstLink); // Replace all
        zip.file("xl/externalLinks/_rels/externalLink1.xml.rels", content);
    }

    const outBuffer = await zip.generateAsync({ type: "arraybuffer" });
    return { name: file.name.replace(sourceYear.slice(2), targetYear.slice(2)), buffer: outBuffer };
}

function copyCell(srcSheet: ExcelJS.Worksheet, dstSheet: ExcelJS.Worksheet, srcAddr: string, dstAddr: string) {
    const srcCell = srcSheet.getCell(srcAddr);
    const dstCell = dstSheet.getCell(dstAddr);

    dstCell.value = srcCell.value;
    dstCell.style = srcCell.style;
}
