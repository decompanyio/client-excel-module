import * as XLSX from "xlsx-js-style";

export type ExcelExportType = "xlsx" | "xls" | "csv" | "txt";
export type ExcelCellStyle = Partial<NonNullable<XLSX.CellObject["s"]>>;

export type ExcelColumn<T> = {
  key: keyof T;
  header: string;
  formatter?: (value: any, row: T) => any;
  style?: ExcelCellStyle;
  wch?: number;
};

export type DownloadExcelOptions<T extends object> = {
  sheetsData: Record<string, T[]>;
  sheetsCol: Record<string, ExcelColumn<T>[]>;
  filename: string;
  type?: ExcelExportType;
  styleOptions?: {
    headerStyle?: ExcelCellStyle;
    bodyStyle?: ExcelCellStyle;
  };
  multiHeadersMap?: Record<string, string[][]>;
};
