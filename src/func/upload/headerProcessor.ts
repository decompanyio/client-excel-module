import * as XLSX from "xlsx-js-style";

/**
 * 단일 헤더를 추출합니다.
 */
export const extractSingleHeader = (
  sheet: XLSX.WorkSheet,
  headerRowIndex: number
): string[] => {
  const headerRows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    range: headerRowIndex,
    blankrows: false,
  }) as any[][];

  return headerRows[0] || [];
};

/**
 * 멀티 헤더를 추출하고 병합합니다.
 */
export const extractMultiHeader = (
  sheet: XLSX.WorkSheet,
  headerRowIndex: number,
  headerRowCount: number
): string[] => {
  const headerRows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    range: headerRowIndex,
    blankrows: false,
  }) as any[][];

  const multiHeaders = headerRows.slice(0, headerRowCount);

  return multiHeaders[0].map((_, colIdx) =>
    multiHeaders
      .map((row) => row[colIdx] ?? "")
      .join(" ")
      .trim()
  );
};

/**
 * 헤더를 추출합니다 (단일/멀티 헤더 자동 판단).
 */
export const extractHeader = (
  sheet: XLSX.WorkSheet,
  headerRowIndex: number,
  headerRowCount: number
): string[] => {
  if (headerRowCount > 1) {
    return extractMultiHeader(sheet, headerRowIndex, headerRowCount);
  } else {
    return extractSingleHeader(sheet, headerRowIndex);
  }
};
