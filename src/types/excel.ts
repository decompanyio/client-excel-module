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

export interface ColumnDefinition<T = any> {
  key: keyof T;
  header: string;
  formatter?: (value: any, row: T) => any;
  wch?: number; // 컬럼 너비
  style?: any; // 셀 스타일
}

export interface StyleOptions {
  headerStyle?: any;
  bodyStyle?: any;
}

export interface UploadExcelOptions {
  sheetName?: string | string[]; // 특정 시트명 또는 시트명 배열
  headerRowIndex?: number; // 헤더가 시작되는 row 인덱스 (기본값: 0)
  headerRowCount?: number; // 헤더 행 개수(멀티 헤더 지원, 기본값: 1)
}

export type DownloadExcelOptions<T extends object> = {
  sheetsData: Record<string, T[]>;
  sheetsCol?: Record<string, ExcelColumn<T>[]>;
  filename?: string;
  type?: ExcelExportType;
  styleOptions?: {
    headerStyle?: ExcelCellStyle;
    bodyStyle?: ExcelCellStyle;
  };
  multiHeadersMap?: Record<string, string[][]>;
};
