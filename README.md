# 使用说明

## 安装与导入

### 安装

```bash
npm i @handsomewolf/excel-export
```

### 导入

``` bash
import { excelExport } from "@handsomewolf/excel-export";
```

## 语法

```ts
excelExport("工作簿名称",工作表数组)
```

### 工作表数组

| 参数 | 类型 | 作用 |
| --- | --- | --- |
| sheetName | string | 工作表名称 |
| sheetHeader | (string\|number)[] | 表头 |
| sheetRows | (string\|number)[]\|{ [key: string]: any, [index: number]: any } | 表头 |
| option | SheetOptions | [可选参数](#可选项)  |

### 可选项

| 参数 | 类型 | 作用 | 默认值 |
| --- | --- | --- | --- |
| sheetStyle | string | 该工作表的样式风格 | 无 |
| accountingColumns | number[] | 列设置为会计专用格式 | 无 |
| decimalPlaces | number | 会计专用格式保留的小数位 | 无 |
| rowFontColor | {rows:number[],defaultColor:string} | 设置某行的字体颜色。rows：数组表示为哪几行设置字体颜色；defaultColor：表示这些行的字体颜色 | FF0000 红色 |

## 示例1

```ts
excelExport("工作簿名称1", [
  {
    sheetName: "工作表名称",
    sheetHeader: ["第1列表头名称", "第2列表头名称", "第3列表头名称"],
    sheetRows: [
      ["第1行1列数据", "第1行2列数据", "第1行3列数据"],
      ["第2行1列数据", "第2行2列数据", "第2行3列数据"],
      ["第3行1列数据", "第3行2列数据", "第3行3列数据"],
    ],
    options: {
      sheetStyle: "green", //excel样式样式风格，目前只有green风格
      accountingColumns: [1, 2], //为1,2列设置为会计专用格式
      decimalPlaces: 2, //accountingColumns（会计专用列）保留几位小数
      rowFontColor: {
        //设置某行的字体颜色
        rows: [1, 2], // 为1,2行设置字体颜色
        defaultColor: "#FF0000", // 可选参数，设置低1，2行的字体颜色为：#FF0000，默认值为FF0000（红色）
      },
      rowBackgroundColors: ["FFF6F4EE", "#EEF8FB"], //每行的背景颜色，该值表示第一行使用#F6F4EE，第二行使用#EEF8FB背景色
    },
  },
]);
```

## 示例2
```ts
const excelHeader = {
  x: "第1列表头名称",
  y: "第2列表头名称",
  z: "第3列表头名称",
};
const excelBody = [
  { x: "第1行1列数据", y: "第1行2列数据", z: "第1行3列数据" },
  { x: "第2行1列数据", y: "第2行2列数据", z: "第2行3列数据" },
  { x: "第3行1列数据", y: "第3行2列数据", z: "第3行3列数据" },
];

excelExport("工作簿名称2", [
  {
    sheetName: "工作表名称",
    sheetHeader: excelHeader,
    sheetRows: excelBody,
    options: {
      sheetStyle: "green", //excel样式样式风格，目前只有green风格
      accountingColumns: [1, 2], //为1,2列设置为会计专用格式
      decimalPlaces: 2, //accountingColumns（会计专用列）保留几位小数
      rowFontColor: {
        //设置某行的字体颜色
        rows: [1, 2], // 为1,2行设置字体颜色
        defaultColor: "#FF0000", // 可选参数，设置低1，2行的字体颜色为：#CCFF00，默认值为FF0000（红色）
      },
      rowBackgroundColors: ["#F6F4EE", "#EEF8FB"], //每行的背景颜色，该值表示第一行使用#F6F4EE，第二行使用#EEF8FB背景色
    },
  },
]);
```
