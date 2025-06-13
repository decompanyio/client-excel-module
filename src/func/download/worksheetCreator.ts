import * as XLSX from "xlsx-js-style";

import type { ColumnDefinition, StyleOptions } from "@/types/excel.ts";
import { convertDataToAOA } from "@/func/download/dataProcessor";
import {
  setColumnWidths,
  setRowHeights,
  applyWorksheetStyles,
} from "@/func/download/worksheetStyler";

/**
 * 단일 워크시트를 생성합니다.
 */
export const createWorksheet = <T extends object>(
  data: T[],
  columns?: ColumnDefinition<T>[],
  multiHeaders?: any[][],
  styleOptions?: StyleOptions,
  fileType: string = "xlsx"
): XLSX.WorkSheet => {
  // 데이터를 2차원 배열로 변환
  const aoa = convertDataToAOA(data, columns, multiHeaders);

  // 워크시트 생성
  const worksheet = XLSX.utils.aoa_to_sheet(aoa);

  // 컬럼 너비 설정
  setColumnWidths(worksheet, columns, aoa[0]?.length);

  // 행 높이 설정
  const headerRows = (multiHeaders?.length ?? 0) + 1;
  setRowHeights(worksheet, aoa.length, headerRows);

  // 스타일 적용
  applyWorksheetStyles(worksheet, columns, styleOptions, headerRows, fileType);

  return worksheet;
};
