import ExcelJS, { type Worksheet } from "exceljs";
import saveAs from "file-saver";
import { greenStyle } from "../style/green";
import { isChinese, isNumber } from "./verify";

export class ExcelHelper {
  private workbook: ExcelJS.Workbook;

  constructor() {
    this.workbook = new ExcelJS.Workbook();
  }

  // 创建工作表
  createSheet(sheetName: string) {
    const sheet = this.workbook.addWorksheet(sheetName);
    return sheet;
  }
  // 插入表头
  insertHeader(sheet: ExcelJS.Worksheet, header: (string | number)[]) {
    const headerRow = sheet.addRow(header);
    headerRow.eachCell((cell) => {
      cell.style = greenStyle.header;
    });
    headerRow.height = 34;
  }

  // 插入数组数据
  insertData(
    sheet: ExcelJS.Worksheet,
    headerRow: (string | number)[],
    bodyRow: SheetRow[],
  ) {
    for (const row of bodyRow) {
      //bodyRow补全
      if (headerRow.length > row.length) {
        for (let index = 0; index < headerRow.length - row.length; index++) {
          row.push("-");
        }
      }
      //undefined、null、“”转换为-
      for (let index = 0; index < row.length; index++) {
        if (
          row[index] === undefined ||
          row[index] === null ||
          row[index] === ""
        ) {
          row[index] = "-";
        }
      }
    }

    sheet.addRows(bodyRow);

    sheet.eachRow((row, rowNumber) => {
      // 排除表头行
      if (rowNumber > 1) {
        row.eachCell((cell) => {
          cell.style = { ...cell.style, ...greenStyle.body };
        });
        row.height = 28;
      } else {
        row.eachCell((cell) => {
          cell.style = { ...greenStyle.header, numFmt: cell.numFmt };
        });
      }
    });
  }

  // 插入对象数据
  // headerData为对象{a:"序号",b:"标题",c:"n内容"} bodyData为对象{b:2,a:1,c:3}，根据headerData的key值进行排序，如果bodyData中没有headerData的key值，则插入-，
  insertObjectData(
    sheet: ExcelJS.Worksheet,
    headerData: { [key: string]: any },
    bodyData: { [key: string]: any }[],
  ) {
    //获取表头的key
    const headerKeys = Object.keys(headerData);
    //获取表头的value
    const headerValues = Object.values(headerData);
    const bodyRow = bodyData.map((item) => {
      const row: (string | number)[] = [];
      for (const key of headerKeys) {
        if (item[key] === undefined) {
          row.push("-");
        } else {
          row.push(item[key]);
        }
      }
      return row;
    });
    // 插入表头
    this.insertHeader(sheet, headerValues);
    // 插入数据
    this.insertData(sheet, headerKeys, bodyRow);
  }

  // 该函数为某一列设置会计专用格式
  setAccountingFormat(
    sheet: Worksheet,
    column: number[],
    decimalPlaces: number | number[] = 2,
  ) {
    // 将字符串转换为数字
    sheet.eachRow((row) => {
      for (const col of column) {
        const cell = row.getCell(col);
        if (typeof cell.value === "string") {
          const numberValue = Number.parseFloat(cell.value);
          if (!Number.isNaN(numberValue)) {
            cell.value = numberValue;
          }
        }
      }
    });
    //
    if (typeof decimalPlaces === "number") {
      const format = `￥#,##0.${"0".repeat(
        decimalPlaces,
      )};￥-#,##0.${"0".repeat(decimalPlaces)};-`;
      for (const col of column) {
        sheet.getColumn(col).numFmt = format;
      }
    } else if (Array.isArray(decimalPlaces)) {
      for (const [index, element] of column.entries()) {
        const format = `￥#,##0.${"0".repeat(
          decimalPlaces[index],
        )};￥-#,##0.${"0".repeat(decimalPlaces[index])};-`;
        sheet.getColumn(element).numFmt = format;
      }
    }
  }

  setSheetStyle(sheet: Worksheet, options: SheetOptions) {
    if (options.accountingColumns) {
      // 设置会计专用格式
      this.setAccountingFormat(sheet, options.accountingColumns);
    }
    if (options.rowFontColor) {
      // 某行设置字体颜色
      this.setRowFontColor(
        sheet,
        options.rowFontColor.rows,
        options.rowFontColor.defaultColor,
      );
    }
    if (options.rowBackgroundColors) {
      // 1某行设置背景颜色
      // this.setRowBackgroundColor(
      //   sheet,
      //   options.rowBackgroundColors.rows,
      //   options.rowBackgroundColors.defaultColor,
      // );
      this.setRowEachBackgroundColor(sheet, options.rowBackgroundColors);
    }
  }

  /**
   * if (sheetData.options) {
    // 4.设置会计专用格式
    if (sheetData.options.accountingColumns.length > 0) {
      excelHelper.setAccountingFormat(
        sheet,
        sheetData.accountingColumns,
        sheetData.decimalPlaces,
      );
    }
    // 9.某行设置字体颜色
    if (sheetData.options.rowFontColor) {
      excelHelper.setRowFontColor(
        sheet,
        sheetData.options.rowFontColor.rows,
        sheetData.options.rowFontColor.defaultColor,
      );
    }
    // 10.某行设置背景颜色
    if (sheetData.options.rowBackgroundColors) {
      excelHelper.setRowEachBackgroundColor(
        sheet,
        sheetData.options.rowBackgroundColors,
      );
    }
  }
   */

  // 导出 Excel 文件
  exportFile(excelName: string) {
    this.workbook.xlsx.writeBuffer().then((buffer) => {
      saveAs(
        new Blob([buffer], { type: "application/octet-stream" }),
        `${excelName}.xlsx`,
      );
    });
  }

  // 自适应列宽度
  autoFitColumn(sheet: Worksheet) {
    for (const column of sheet.columns) {
      let maxPixelWidth = 0;
      column.eachCell!({ includeEmpty: true }, (cell) => {
        let pixelWidth = 0;
        for (const char of cell.text) {
          pixelWidth += isChinese(char) ? 3 : isNumber(char) ? 2.7 : 1.5; // 如果字符是中文，宽度为2，否则为1
        }
        if (pixelWidth > maxPixelWidth) {
          maxPixelWidth = pixelWidth;
        }
      });
      column.width = maxPixelWidth < 10 ? 10 : maxPixelWidth; // 设置最小宽度为10
    }
  }

  // 为某一行设置字体颜色，参数为数组表示哪行
  setRowFontColor(sheet: Worksheet, row: number[], color: string = "FF0000") {
    for (const rowNumber of row) {
      sheet.getRow(rowNumber + 1).eachCell((cell) => {
        cell.font = { ...cell.font, color: { argb: color.replace("#", "") } };
      });
    }
  }

  //为某一行设置背景颜色，参数为数组表示哪行
  setRowBackgroundColor(
    sheet: Worksheet,
    row: number[],
    color: string = "f8fffd",
  ) {
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

  //逐行设置背景色
  setRowEachBackgroundColor(sheet: Worksheet, color: string[]) {
    sheet.eachRow((row, rowNumber) => {
      // 排除表头行
      if (rowNumber > 1) {
        row.eachCell((cell) => {
          // 设置背景颜色
          cell.style.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: color[rowNumber - 2] },
          };
        });
      }
    });
  }
}
