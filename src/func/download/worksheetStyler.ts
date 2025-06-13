import * as XLSX from "xlsx-js-style";

import type { ColumnDefinition, StyleOptions } from "@/types/excel.ts";

/**
 * 워크시트의 컬럼 너비를 설정합니다.
 */
export const setColumnWidths = <T>(
  worksheet: XLSX.WorkSheet,
  columns?: ColumnDefinition<T>[],
  aoaLength?: number
): void => {
  if (Array.isArray(columns) && columns.length > 0) {
    worksheet["!cols"] = columns.map((col) => ({ wch: col.wch ?? 20 }));
  } else {
    const colLen = aoaLength ?? 0;
    worksheet["!cols"] = Array.from({ length: colLen }, () => ({ wch: 20 }));
  }
};

/**
 * 워크시트의 행 높이를 설정합니다.
 */
export const setRowHeights = (
  worksheet: XLSX.WorkSheet,
  totalRows: number,
  headerRows: number
): void => {
  worksheet["!rows"] = [];
  for (let i = 0; i < totalRows; i++) {
    worksheet["!rows"].push({
      hpt: i < headerRows ? 32 : 24,
    });
  }
};

/**
 * 워크시트에 스타일을 적용합니다.
 */
export const applyWorksheetStyles = <T>(
  worksheet: XLSX.WorkSheet,
  columns?: ColumnDefinition<T>[],
  styleOptions?: StyleOptions,
  headerRows: number = 1,
  fileType: string = "xlsx"
): void => {
  if (fileType !== "xlsx" && fileType !== "xls") return;

  const range = XLSX.utils.decode_range(worksheet["!ref"] as string);

  // 헤더 스타일 적용
  applyHeaderStyles(worksheet, range, headerRows, columns, styleOptions);

  // 바디 스타일 적용
  applyBodyStyles(worksheet, range, headerRows, styleOptions);
};

/**
 * 헤더 스타일을 적용합니다.
 */
const applyHeaderStyles = <T>(
  worksheet: XLSX.WorkSheet,
  range: XLSX.Range,
  headerRows: number,
  columns?: ColumnDefinition<T>[],
  styleOptions?: StyleOptions
): void => {
  for (let R = 0; R < headerRows; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      if (!worksheet[cellAddress]) continue;

      // 기본 헤더 스타일
      let cellStyle = {
        font: { bold: true, color: { rgb: "FFFFFF" }, sz: 16 },
        fill: { fgColor: { rgb: "4F81BD" } },
        alignment: { horizontal: "center", vertical: "center" },
        ...styleOptions?.headerStyle,
      };

      // 컬럼별 스타일 적용 (마지막 헤더 행)
      if (R === headerRows - 1 && Array.isArray(columns) && columns[C]?.style) {
        cellStyle = { ...cellStyle, ...columns[C].style };
      }

      worksheet[cellAddress].s = cellStyle;
    }
  }
};

/**
 * 바디 스타일을 적용합니다.
 */
const applyBodyStyles = (
  worksheet: XLSX.WorkSheet,
  range: XLSX.Range,
  headerRows: number,
  styleOptions?: StyleOptions
): void => {
  for (let R = headerRows; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      if (!worksheet[cellAddress]) continue;

      worksheet[cellAddress].s = {
        font: { sz: 14 },
        alignment: { vertical: "center" },
        ...styleOptions?.bodyStyle,
      };
    }
  }
};
