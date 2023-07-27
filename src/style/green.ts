import { type Style } from "exceljs";

export const greenStyle: { header: Partial<Style>; body: Partial<Style> } = {
  header: {
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF12BA83" },
    },
    font: {
      name: "微软雅黑",
      size: 12,
      color: { argb: "FFFFFFFF" },
      bold: true,
    },
    border: {
      top: { style: "thin", color: { argb: "FF28C593" } },
      left: { style: "thin", color: { argb: "FF28C593" } },
      bottom: { style: "thin", color: { argb: "FF28C593" } },
      right: { style: "thin", color: { argb: "FF28C593" } },
    },
    alignment: {
      vertical: "middle",
      horizontal: "center",
    },
  },
  body: {
    font: {
      name: "微软雅黑",
      size: 12,
      color: { argb: "FF333333" },
    },
    border: {
      top: { style: "thin", color: { argb: "FFD6D8D7" } },
      left: { style: "thin", color: { argb: "FFD6D8D7" } },
      bottom: { style: "thin", color: { argb: "FFD6D8D7" } },
      right: { style: "thin", color: { argb: "FFD6D8D7" } },
    },
    alignment: {
      vertical: "middle",
      horizontal: "center",
    },
  },
};
