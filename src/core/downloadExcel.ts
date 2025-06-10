import * as XLSX from "xlsx-js-style";
import type { DownloadExcelOptions } from "@/types/excel.ts";

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
    sheetsData, // 각 시트별 데이터
    sheetsCol, // 각 시트별 컬럼 정의 (헤더, 포맷터, 스타일 등)
    filename = "sheet", // 저장 파일명
    type = "xlsx", // 파일 타입 (기본값: xlsx)
    styleOptions, // 전체 헤더/바디 스타일
    multiHeadersMap, // 멀티 헤더(2차원 배열) 정의
  } = options;

  const workbook = XLSX.utils.book_new();

  // 각 시트별로 처리
  for (const [sheetName, data] of Object.entries(sheetsData)) {
    const columns = sheetsCol?.[sheetName];
    let aoa: any[][] = [];
    const multiHeaders = multiHeadersMap?.[sheetName];

    // 멀티 헤더가 있으면 먼저 추가
    if (multiHeaders && multiHeaders.length > 0) {
      aoa = [...multiHeaders];
    }

    if (Array.isArray(columns) && columns.length > 0) {
      // 컬럼 헤더 추가
      aoa.push(columns.map((col) => col.header));

      // 데이터 행 추가 (포맷터 적용)
      if (data && data.length > 0) {
        for (const row of data) {
          aoa.push(
            columns.map((col) => {
              try {
                let value = (row as any)[col.key];
                if (col.formatter) value = col.formatter(value, row);
                return value;
              } catch (error) {
                console.error(
                  `Error formatting column ${String(col.key)}:`,
                  error
                );
                return (row as any)[col.key] || "";
              }
            })
          );
        }
      }
    } else {
      // 컬럼 정의가 없을 때: 데이터의 키를 자동으로 사용
      const keys = Object.keys(data[0] ?? {});
      aoa.push(keys);
      if (data && data.length > 0) {
        for (const row of data) {
          aoa.push(keys.map((k) => (row as Record<string, any>)[k]));
        }
      }
    }

    // 시트 생성
    const worksheet = XLSX.utils.aoa_to_sheet(aoa);

    // 컬럼별 너비 지정 (sheetsCol의 width 우선, 없으면 기본값 20)
    if (Array.isArray(columns) && columns.length > 0) {
      worksheet["!cols"] = columns.map((col) => ({ wch: col.wch ?? 20 }));
    } else {
      const colLen = aoa[0]?.length ?? 0;
      worksheet["!cols"] = Array.from({ length: colLen }, () => ({ wch: 20 }));
    }

    // 행 높이 지정 (기본: 헤더 32, 바디 24)
    worksheet["!rows"] = [];
    const headerRows = (multiHeaders?.length ?? 0) + 1;
    for (let i = 0; i < aoa.length; ++i) {
      worksheet["!rows"].push({
        hpt: i < headerRows ? 32 : 24,
      });
    }

    // 스타일 적용 (헤더/바디)
    if (type === "xlsx" || type === "xls") {
      const range = XLSX.utils.decode_range(worksheet["!ref"] as string);
      const headerRows = (multiHeaders?.length ?? 0) + 1;

      // 헤더 스타일
      for (let R = 0; R < headerRows; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
          if (!worksheet[cellAddress]) continue;

          // 기본 헤더 스타일 + 전체 헤더 스타일 + 컬럼별 스타일(마지막 헤더 행)
          let cellStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 16 },
            fill: { fgColor: { rgb: "4F81BD" } },
            alignment: { horizontal: "center", vertical: "center" },
            ...styleOptions?.headerStyle,
          };

          if (
            R === headerRows - 1 &&
            Array.isArray(columns) &&
            columns[C]?.style
          ) {
            cellStyle = { ...cellStyle, ...columns[C].style };
          }

          worksheet[cellAddress].s = cellStyle;
        }
      }

      // 바디(데이터) 스타일
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

    // 시트 추가
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  }

  // 파일 다운로드 처리
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
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
  } else {
    XLSX.writeFile(workbook, `${filename}.${type}`);
  }
};
