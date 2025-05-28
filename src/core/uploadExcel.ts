import * as XLSX from "xlsx";

/**
 * 엑셀 파일을 읽어 여러 시트의 데이터를 JSON 배열로 반환합니다.
 * 멀티 헤더(여러 줄 헤더)도 지원하며, 원하는 행을 헤더로 지정할 수 있습니다.
 *
 * @template T 반환할 객체 타입
 * @param {File} file - 읽을 엑셀 파일 객체
 * @param {Object} [options] - 옵션 객체
 * @param {string|string[]} [options.sheetName] - 읽을 시트명 또는 시트명 배열 (기본값: 모든 시트)
 * @param {number} [options.headerRowIndex] - 헤더가 시작되는 행 인덱스 (기본값: 0)
 * @param {number} [options.headerRowCount] - 헤더 행 개수(멀티 헤더 지원, 기본값: 1)
 * @returns {Promise<Record<string, T[]>>} 시트별 데이터 객체 반환
 */
export const uploadExcel = <T extends object>(
  file: File,
  options?: {
    sheetName?: string | string[]; // 특정 시트명 또는 시트명 배열
    headerRowIndex?: number; // 헤더가 시작되는 row 인덱스 (기본값: 0)
    headerRowCount?: number; // 헤더 행 개수(멀티 헤더 지원, 기본값: 1)
  }
): Promise<Record<string, T[]>> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });

        // 시트명 배열로 변환 (없으면 전체 시트)
        let sheetNames: string[] = [];
        if (!options?.sheetName) {
          sheetNames = workbook.SheetNames;
        } else if (typeof options.sheetName === "string") {
          sheetNames = [options.sheetName];
        } else {
          sheetNames = options.sheetName;
        }

        const headerRowIndex = options?.headerRowIndex ?? 0;
        const headerRowCount = options?.headerRowCount ?? 1;

        const result: Record<string, T[]> = {};

        for (const sheetName of sheetNames) {
          const sheet = workbook.Sheets[sheetName];
          if (!sheet) continue;

          // 멀티 헤더 지원: 여러 줄 헤더 병합
          let header: string[] = [];
          if (headerRowCount > 1) {
            const headerRows = XLSX.utils.sheet_to_json(sheet, {
              header: 1,
              range: headerRowIndex,
              blankrows: false,
            }) as any[][];
            const multiHeaders = headerRows.slice(0, headerRowCount);
            header = multiHeaders[0].map((_, colIdx) =>
              multiHeaders
                .map((row) => row[colIdx] ?? "")
                .join(" ")
                .trim()
            );
          } else {
            const headerRows = XLSX.utils.sheet_to_json(sheet, {
              header: 1,
              range: headerRowIndex,
              blankrows: false,
            }) as any[][];
            header = headerRows[0] || [];
          }

          // 데이터 추출 (헤더 아래부터)
          const dataStartRow = headerRowIndex + headerRowCount;
          const json = XLSX.utils.sheet_to_json<T>(sheet, {
            header: header,
            range: dataStartRow,
            defval: "",
          });

          result[sheetName] = json;
        }

        resolve(result);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = () => reject(new Error("파일 읽기 실패"));
    reader.readAsArrayBuffer(file);
  });
};
