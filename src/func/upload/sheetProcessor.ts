import * as XLSX from "xlsx-js-style";

import { extractHeader } from "@/func/upload/headerProcessor.ts";

/**
 * 단일 시트의 데이터를 추출합니다.
 */
export const extractSheetData = <T extends object>(
  sheet: XLSX.WorkSheet,
  headerRowIndex: number,
  headerRowCount: number
): T[] => {
  // 헤더 추출
  const header = extractHeader(sheet, headerRowIndex, headerRowCount);

  // 데이터 추출 (헤더 아래부터)
  const dataStartRow = headerRowIndex + headerRowCount;
  const json = XLSX.utils.sheet_to_json<T>(sheet, {
    header: header,
    range: dataStartRow,
    defval: "",
  });

  return json;
};

/**
 * 여러 시트의 데이터를 추출합니다.
 */
export const extractSheetsData = <T extends object>(
  workbook: XLSX.WorkBook,
  sheetNames: string[],
  headerRowIndex: number,
  headerRowCount: number
): Record<string, T[]> => {
  const result: Record<string, T[]> = {};

  for (const sheetName of sheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;

    result[sheetName] = extractSheetData<T>(
      sheet,
      headerRowIndex,
      headerRowCount
    );
  }

  return result;
};
