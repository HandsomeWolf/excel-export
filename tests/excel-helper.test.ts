import { beforeEach, describe, expect, it } from "vitest";
import { ExcelHelper } from "../src/utils/excel-helper";

describe("ExcelHelper", () => {
  let excelHelper: ExcelHelper;

  beforeEach(() => {
    excelHelper = new ExcelHelper();
  });

  describe("createSheet", () => {
    it("should create a new sheet with the given name", () => {
      const sheetName = "Test Sheet";
      const sheet = excelHelper.createSheet(sheetName);
      expect(sheet.name).to.equal(sheetName);
    });
  });

  describe("insertHeader", () => {
    it("should insert a header row into the sheet", () => {
      const sheet = excelHelper.createSheet("Test Sheet");
      const headers = ["Name", "Age", "Email"];
      excelHelper.insertHeader(sheet, headers);
      expect(sheet.getRow(1).values).to.deep.equal(headers);
    });
  });

  describe("insertData", () => {
    it("should insert data into the sheet", () => {
      const sheet = excelHelper.createSheet("Test Sheet");
      const data = [["John Doe", 30, "john.doe@example.com"]];
      excelHelper.insertData(sheet, data);
      expect(sheet.getRow(2).values).to.deep.equal(data[0]);
    });
  });
});
