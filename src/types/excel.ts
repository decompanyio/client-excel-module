import * as XLSX from "xlsx-js-style";

export type ExcelExportType = "xlsx" | "xls" | "csv" | "txt";
export type ExcelCellStyle = Partial<NonNullable<XLSX.CellObject["s"]>>;
export type DownloadExcelOptions<T extends object> = {
  sheetsData: Record<string, T[]>;
  filename: string;
  type?: ExcelExportType;
  styleOptions?: {
    headerStyle?: ExcelCellStyle;
    bodyStyle?: ExcelCellStyle;
  };
  multiHeadersMap?: Record<string, string[][]>;
};
