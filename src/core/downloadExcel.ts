// utils/excel/index.ts
import * as XLSX from "xlsx-js-style";

import type { DownloadExcelOptions } from "@/types/excel.ts";
import { createWorksheet } from "@/func/download/worksheetCreator.ts";
import {
  downloadTextFile,
  downloadExcelFile,
} from "@/func/download/fileDownloader.ts";

/**
 * 여러 워크시트의 데이터를 엑셀 파일로 변환하여 다운로드합니다.
 * 컬럼 정의, 포맷팅, 멀티 헤더, 셀 스타일 커스텀, 컬럼별 스타일 지원.
 *
 * @template T - 데이터 객체 타입
 * @param options - 다운로드 옵션
 */
export const downloadExcel = <T extends object>(
  options: DownloadExcelOptions<T>
): void => {
  const {
    sheetsData,
    sheetsCol,
    filename = "sheet",
    type = "xlsx",
    styleOptions,
    multiHeadersMap,
  } = options;

  const workbook = XLSX.utils.book_new();

  // 각 시트별로 처리
  for (const [sheetName, data] of Object.entries(sheetsData)) {
    const columns = sheetsCol?.[sheetName];
    const multiHeaders = multiHeadersMap?.[sheetName];

    // 워크시트 생성
    const worksheet = createWorksheet(
      data,
      columns,
      multiHeaders,
      styleOptions,
      type
    );

    // 워크북에 시트 추가
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  }

  // 파일 다운로드
  if (type === "csv" || type === "txt") {
    downloadTextFile(workbook, filename, type);
  } else {
    downloadExcelFile(workbook, filename, type);
  }
};
