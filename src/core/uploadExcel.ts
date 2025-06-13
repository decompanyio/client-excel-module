import * as XLSX from "xlsx-js-style";

import type { UploadExcelOptions } from "@/types/excel.ts";
import { readFileAsArrayBuffer } from "@/func/upload/fileReader.ts";
import { processWorkbook } from "@/func/upload/workbookProcessor.ts";
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
export const uploadExcel = async <T extends object>(
  file: File,
  options?: UploadExcelOptions
): Promise<Record<string, T[]>> => {
  try {
    // 파일을 ArrayBuffer로 읽기
    const arrayBuffer = await readFileAsArrayBuffer(file);

    // 워크북 생성
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: "array" });

    // 워크북 처리 및 데이터 추출
    return processWorkbook<T>(workbook, options);
  } catch (error) {
    throw new Error(`엑셀 파일 처리 중 오류가 발생했습니다: ${error}`);
  }
};
