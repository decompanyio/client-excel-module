import * as XLSX from "xlsx-js-style";

import type { UploadExcelOptions } from "@/types/excel.ts";
import { resolveSheetNames } from "@/func/upload/sheetNameResolver.ts";
import { extractSheetsData } from "@/func/upload/sheetProcessor.ts";

/**
 * 워크북을 처리하여 데이터를 추출합니다.
 */
export const processWorkbook = <T extends object>(
  workbook: XLSX.WorkBook,
  options?: UploadExcelOptions
): Record<string, T[]> => {
  const headerRowIndex = options?.headerRowIndex ?? 0;
  const headerRowCount = options?.headerRowCount ?? 1;

  // 처리할 시트명 결정
  const sheetNames = resolveSheetNames(workbook, options?.sheetName);

  // 시트별 데이터 추출
  return extractSheetsData<T>(
    workbook,
    sheetNames,
    headerRowIndex,
    headerRowCount
  );
};
