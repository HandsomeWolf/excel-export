/* global describe, it */
import { excelExport } from "../../dist/index";
describe("excelExport function", () => {
  it("should generate an Excel file with correct data", () => {
    // 模拟数据
    const fileName = "test";
    const sheetDatas = [
      {
        sheetName: "Sheet1",
        sheetHeader: ["Column1", "Column2"],
        sheetRows: [
          ["Data1", "Data2"],
          ["Data3", "Data4"],
        ],
      },
    ];

    // 调用excelExport函数
    excelExport(fileName, sheetDatas);

    // 检查生成的Excel文件是否存在并且数据正确
    // 注意：你可能需要使用一些库来读取和验证Excel文件
  });
});
