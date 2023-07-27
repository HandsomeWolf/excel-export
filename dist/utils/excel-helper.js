"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelHelper = void 0;
const exceljs_1 = __importDefault(require("exceljs"));
const file_saver_1 = __importDefault(require("file-saver"));
const green_1 = require("../style/green");
const verify_1 = require("./verify");
class ExcelHelper {
    constructor() {
        this.workbook = new exceljs_1.default.Workbook();
    }
    // 创建工作表
    createSheet(sheetName) {
        const sheet = this.workbook.addWorksheet(sheetName);
        return sheet;
    }
    // 插入表头
    insertHeader(sheet, header) {
        const headerRow = sheet.addRow(header);
        headerRow.eachCell((cell) => {
            cell.style = green_1.greenStyle.header;
        });
        headerRow.height = 34;
    }
    // 插入数组数据
    insertData(sheet, row) {
        sheet.addRows(row);
        sheet.eachRow((row, rowNumber) => {
            // 排除表头行
            if (rowNumber > 1) {
                row.eachCell((cell) => {
                    cell.style = green_1.greenStyle.body;
                });
                row.height = 28;
            }
            else {
                row.eachCell((cell) => {
                    cell.style = green_1.greenStyle.header;
                });
            }
        });
    }
    // 该函数为某一列设置会计专用格式
    setAccountingFormat(sheet, column) {
        sheet.eachRow((row, rowNumber) => {
            for (const col of column) {
                const cell = row.getCell(col);
                cell.numFmt = "￥#,##0.00;￥(#,##0.00);-";
            }
        });
    }
    // 导出 Excel 文件
    exportFile(excelName) {
        this.workbook.xlsx.writeBuffer().then((buffer) => {
            (0, file_saver_1.default)(new Blob([buffer], { type: "application/octet-stream" }), `${excelName}.xlsx`);
        });
    }
    // // 创建工作表并填充数据
    // createSheetAndFillData(
    //   sheetName: string,
    //   sheetHeader: string[],
    //   sheetData: { [key: string]: any }[],
    // ) {
    //   const sheet = this.workbook.addWorksheet(sheetName);
    //   // 创建表头
    //   const headers = ["fapiao", "hetong", "jine"].map((key, index) => ({
    //     header: ["发票", "合同", "金额"][index],
    //     key,
    //   }));
    //   sheet.columns = headers;
    //   // 填充数据并设置样式
    //   for (const item of sheetData) {
    //     const row = sheet.addRow([item.fapiao, item.hetong, item.jine]);
    //     row.eachCell((cell) => {
    //       cell.style = {
    //         font: { color: { argb: item.other.color.replace("#", "") } },
    //         fill: {
    //           type: "pattern",
    //           pattern: "solid",
    //           fgColor: { argb: item.other.background.replace("#", "") },
    //         },
    //       };
    //     });
    //     // 设置金额列为特殊格式
    //     const jineCell = row.getCell(3);
    //     jineCell.numFmt = "¥#,##0.00";
    //   }
    //   // 设置列宽度自适应
    //   for (const column of sheet.columns) {
    //     let maxLength = 0;
    //     column.eachCell!({ includeEmpty: true }, (cell) => {
    //       const cellLength = cell.text?.length || 0;
    //       if (cellLength > maxLength) {
    //         maxLength = cellLength;
    //       }
    //     });
    //     column.width = maxLength < 10 ? 10 : maxLength;
    //   }
    // }
    // 自适应列宽度
    autoFitColumn(sheet) {
        for (const column of sheet.columns) {
            let maxPixelWidth = 0;
            column.eachCell({ includeEmpty: true }, (cell) => {
                let pixelWidth = 0;
                for (const char of cell.text) {
                    pixelWidth += (0, verify_1.isChinese)(char) ? 3 : (0, verify_1.isNumber)(char) ? 2.7 : 1.5; // 如果字符是中文，宽度为2，否则为1
                }
                if (pixelWidth > maxPixelWidth) {
                    maxPixelWidth = pixelWidth;
                }
            });
            column.width = maxPixelWidth < 10 ? 10 : maxPixelWidth; // 设置最小宽度为10
        }
    }
    // 为某一行设置字体颜色，参数为数组表示哪行
    setRowFontColor(sheetName, row, color) {
        const sheet = this.workbook.getWorksheet(sheetName);
        for (const rowNumber of row) {
            sheet.getRow(rowNumber).eachCell((cell) => {
                cell.font.color = { argb: color.replace("#", "") };
            });
        }
    }
    //为某一行设置背景颜色，参数为数组表示哪行
    setRowBackgroundColor(sheetName, row, color) {
        const sheet = this.workbook.getWorksheet(sheetName);
        for (const rowNumber of row) {
            sheet.getRow(rowNumber).eachCell((cell) => {
                cell.fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: { argb: color.replace("#", "") },
                };
            });
        }
    }
}
exports.ExcelHelper = ExcelHelper;
