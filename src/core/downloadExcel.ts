import * as XLSX from "xlsx-js-style";

export type ExcelExportType = "xlsx" | "xls" | "csv" | "txt";
export type ExcelCellStyle = Partial<NonNullable<XLSX.CellObject["s"]>>;

/**
 * 여러 워크시트의 데이터를 엑셀 파일로 변환하여 다운로드합니다.
 * 멀티 헤더, 헤더/바디 스타일 커스텀을 지원합니다.
 *
 * @template T 내보낼 객체 타입
 * @param {Record<string, T[]>} sheetsData - 시트별 데이터 객체 (key: 시트명, value: 데이터 배열)
 * @param {string} filename - 저장할 파일명(확장자 제외)
 * @param {ExcelExportType} [type="xlsx"] - 저장할 파일 형식(xlsx, xls, csv, txt)
 * @param {Object} [styleOptions] - 스타일 옵션
 * @param {ExcelCellStyle} [styleOptions.headerStyle] - 헤더 셀 스타일
 * @param {ExcelCellStyle} [styleOptions.bodyStyle] - 바디 셀 스타일
 * @param {Record<string, string[][]>} [multiHeadersMap] - 시트별 멀티 헤더(2차원 배열)
 */
export const downloadExcel = <T extends object>(
  sheetsData: Record<string, T[]>,
  filename: string,
  type: ExcelExportType = "xlsx",
  styleOptions?: {
    headerStyle?: ExcelCellStyle;
    bodyStyle?: ExcelCellStyle;
  },
  multiHeadersMap?: Record<string, string[][]>
): void => {
  const workbook = XLSX.utils.book_new();

  Object.entries(sheetsData).forEach(([sheetName, data]) => {
    let aoa: any[][] = [];
    const multiHeaders = multiHeadersMap?.[sheetName];

    // 멀티 헤더가 있으면 헤더를 먼저 추가하고, 데이터는 값만 추출해서 추가
    if (multiHeaders && multiHeaders.length > 0) {
      aoa = [...multiHeaders];
      const dataRows = data.map((row) => Object.values(row));
      aoa.push(...dataRows);
    } else {
      // 멀티 헤더가 없으면 첫 줄에 키를 헤더로 사용
      aoa = [
        Object.keys(data[0] ?? {}),
        ...data.map((row) => Object.values(row)),
      ];
    }

    // 2차원 배열을 시트로 변환
    const worksheet = XLSX.utils.aoa_to_sheet(aoa);

    // 스타일 적용 (엑셀 형식일 때만)
    if (type === "xlsx" || type === "xls") {
      const range = XLSX.utils.decode_range(worksheet["!ref"] as string);
      const headerRows = multiHeaders ? multiHeaders.length : 1;
      // 헤더 스타일 적용
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
      // 바디(데이터) 스타일 적용
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

  // csv, txt는 첫 번째 시트만 저장
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
