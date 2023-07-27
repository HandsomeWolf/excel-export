type SheetStyle = "style1" | "style2";
type SheetRow = (string | number)[] | { [key: string]: any, [index: number]: any };
interface RowFontColor {
  colors?: { [key: number]: string };
  rows: number[];
  defaultColor?: string;
}
interface RowBackgroundColor {
  colors?: { [key: number]: string };
  rows: number[];
  defaultColor?: string;
}

interface SheetOptions {
  /**
     * 工作表样式
     */
  sheetStyle?: SheetStyle;
  /**
   * 会计专用格式列
   */
  accountingColumns?: number[],
  /**
   * 会计专用格式小数位数
   */
  decimalPlaces?:number|number[]

  /**
   * 指定行字体颜色
   */
  rowFontColor?: RowFontColor;

  /**
   * 每一行背景颜色
   */
  rowBackgroundColors?: string[];
}

interface SheetParameters{
  /**
   * 工作表名称
   */
  sheetName: string;
  /**
   * 工作表表头
   */
  sheetHeader: (string | number)[];
  /**
   * 工作表数据
   */
  sheetRows: SheetRow[];
  /**
   * 工作表样式
   */
  options?: SheetOptions

}
