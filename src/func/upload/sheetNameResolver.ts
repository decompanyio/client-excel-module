import * as XLSX from "xlsx-js-style";

/**
 * 옵션에 따라 처리할 시트명 목록을 결정합니다.
 */
export const resolveSheetNames = (
  workbook: XLSX.WorkBook,
  sheetNameOption?: string | string[]
): string[] => {
  if (!sheetNameOption) {
    return workbook.SheetNames;
  }

  if (typeof sheetNameOption === "string") {
    return [sheetNameOption];
  }

  return sheetNameOption;
};
