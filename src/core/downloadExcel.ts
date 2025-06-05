import * as XLSX from "xlsx-js-style";

import type { DownloadExcelOptions } from "@/types/excel.ts";

/**
 * 여러 워크시트의 데이터를 엑셀 파일로 변환하여 다운로드합니다.
 * 멀티 헤더, 헤더/바디 스타일 커스텀을 지원합니다.
 *
 * @template T 내보낼 객체 타입
 * @param {DownloadExcelOptions<T>} options - 엑셀 다운로드 옵션 객체
 */
export const downloadExcel = <T extends object>(
  options: DownloadExcelOptions<T>
): void => {
  const {
    sheetsData,
    filename,
    type = "xlsx",
    styleOptions,
    multiHeadersMap,
  } = options;

  const workbook = XLSX.utils.book_new();

  Object.entries(sheetsData).forEach(([sheetName, data]) => {
    let aoa: any[][] = [];
    const multiHeaders = multiHeadersMap?.[sheetName];

    if (multiHeaders && multiHeaders.length > 0) {
      aoa = [...multiHeaders];
      const dataRows = data.map((row) => Object.values(row));
      aoa.push(...dataRows);
    } else {
      aoa = [
        Object.keys(data[0] ?? {}),
        ...data.map((row) => Object.values(row)),
      ];
    }

    const worksheet = XLSX.utils.aoa_to_sheet(aoa);

    if (type === "xlsx" || type === "xls") {
      const range = XLSX.utils.decode_range(worksheet["!ref"] as string);
      const headerRows = multiHeaders ? multiHeaders.length : 1;
      for (let R = 0; R < headerRows; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
          if (!worksheet[cellAddress]) continue;
          worksheet[cellAddress].s = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 16 },
            fill: { fgColor: { rgb: "4F81BD" } },
            alignment: { horizontal: "center", vertical: "center" },
            ...styleOptions?.headerStyle,
          };
        }
      }
      for (let R = headerRows; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
          if (!worksheet[cellAddress]) continue;
          worksheet[cellAddress].s = {
            font: { sz: 14 },
            alignment: { vertical: "center" },
            ...styleOptions?.bodyStyle,
          };
        }
      }
    }

    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  });

  if (type === "csv" || type === "txt") {
    const firstSheet = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheet];
    const output = XLSX.utils.sheet_to_csv(worksheet, {
      FS: type === "csv" ? "," : "\t",
    });
    const blob = new Blob([output], {
      type: "text/plain;charset=utf-8;",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${filename}.${type}`;
    a.click();
    URL.revokeObjectURL(a.href);
  } else {
    XLSX.writeFile(workbook, `${filename}.${type}`);
  }
};
