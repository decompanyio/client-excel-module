import * as XLSX from "xlsx";

/**
 * 시트에서 헤더(첫 번째 행의 값) 목록을 추출합니다.
 *
 * @param {XLSX.WorkSheet} sheet - 엑셀 시트 객체
 * @returns {string[]} 헤더 문자열 배열
 */
export const getHeaders = (
  sheet: XLSX.WorkSheet,
  headerRowIndex: number = 0
): string[] => {
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    range: headerRowIndex,
    blankrows: false,
  }) as any[][];
  return rows[0] || [];
};

/**
 * 시트에서 데이터 행을 JSON 배열로 반환합니다.
 *
 * @template T 반환할 객체 타입
 * @param {XLSX.WorkSheet} sheet - 엑셀 시트 객체
 * @returns {T[]} 데이터 행 객체 배열
 */
export const getRows = <T extends object>(
  sheet: XLSX.WorkSheet,
  headerRowIndex: number = 0 // 헤더가 있는 행 인덱스 지정
): T[] => {
  return XLSX.utils.sheet_to_json<T>(sheet, {
    header: 1, // 2차원 배열로 반환
    range: headerRowIndex,
    defval: "",
  });
};

/**
 * 시트에서 각 컬럼별로 값을 배열로 반환합니다.
 *
 * @template T 반환할 객체 타입
 * @param {XLSX.WorkSheet} sheet - 엑셀 시트 객체
 * @returns {Record<string, any[]>} 컬럼명별 값 배열 객체
 */
export const getColumns = (
  sheet: XLSX.WorkSheet,
  headerRowIndex: number = 0
): Record<string, any[]> => {
  const headers = getHeaders(sheet, headerRowIndex);
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    range: headerRowIndex + 1,
    blankrows: false,
    defval: "",
  }) as any[][];
  const columns: Record<string, any[]> = {};
  headers.forEach((header, idx) => {
    columns[header] = rows.map((row) => row[idx]);
  });
  return columns;
};

/**
 * 시트에서 헤더, 행, 컬럼 정보를 모두 반환합니다.
 *
 * @template T 반환할 객체 타입
 * @param {XLSX.WorkSheet} sheet - 엑셀 시트 객체
 * @returns {{ headers: string[], rows: T[], columns: Record<string, any[]> }} 시트 정보 객체
 */
export const getSheetInfo = (sheet: XLSX.WorkSheet) => {
  const headers = getHeaders(sheet, 0);
  const rows = getRows(sheet, 0);
  const columns = getColumns(sheet, 0);

  return {
    headers,
    rows,
    columns,
  };
};
