import ExcelJS, { type Worksheet } from "exceljs";
export declare class ExcelHelper {
    private workbook;
    constructor();
    createSheet(sheetName: string): ExcelJS.Worksheet;
    insertHeader(sheet: ExcelJS.Worksheet, header: (string | number)[]): void;
    insertData(sheet: ExcelJS.Worksheet, row: SheetRow[]): void;
    setAccountingFormat(sheet: ExcelJS.Worksheet, column: number[]): void;
    exportFile(excelName: string): void;
    autoFitColumn(sheet: Worksheet): void;
    setRowFontColor(sheetName: string, row: number[], color: string): void;
    setRowBackgroundColor(sheetName: string, row: number[], color: string): void;
}
