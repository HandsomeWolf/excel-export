import { type Worksheet } from "exceljs";
import { ExcelHelper } from "./utils/excel-helper";

export const excelExport = (
  fileName: string,
  sheetDatas: SheetParameters[],
) => {
  // 1.创建工作簿
  const excelHelper = new ExcelHelper();

  for (const sheetData of sheetDatas) {
    // 2.创建工作表
    const sheet = excelHelper.createSheet(sheetData.sheetName);

    // 3.向工作表插入数据
    if (Array.isArray(sheetData.sheetRows[0])) {
      insertArrayData(sheet, sheetData, excelHelper);
    } else {
      inertObjectData(sheet, sheetData, excelHelper);
    }
    // 6.自适应列宽
    excelHelper.autoFitColumn(sheet);
  }

  // 8.导出文件
  excelHelper.exportFile(fileName);
};

// 如果是数组就调用这个方法
const insertArrayData = (
  sheet: Worksheet,
  sheetData: SheetParameters,
  excelHelper: ExcelHelper,
) => {
  // 3.插入表头
  excelHelper.insertHeader(sheet, sheetData.sheetHeader);

  // 5.插入数据
  excelHelper.insertData(sheet, sheetData.sheetHeader, sheetData.sheetRows);
  // 4.设置样式
  if (sheetData.options) {
    excelHelper.setSheetStyle(sheet, sheetData.options);
  }
};

// 如果是对象就调用这个方法
const inertObjectData = (
  sheet: Worksheet,
  sheetData: SheetParameters,
  excelHelper: ExcelHelper,
) => {
  // 插入表头 与 数据
  excelHelper.insertObjectData(
    sheet,
    sheetData.sheetHeader,
    sheetData.sheetRows,
  );
  // 设置样式
  if (sheetData.options) {
    excelHelper.setSheetStyle(sheet, sheetData.options);
  }
};
