import type { ColumnDefinition } from "@/types/excel.ts";

/**
 * 데이터를 2차원 배열로 변환합니다.
 */
export const convertDataToAOA = <T extends object>(
  data: T[],
  columns?: ColumnDefinition<T>[],
  multiHeaders?: any[][]
): any[][] => {
  let aoa: any[][] = [];

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

  return aoa;
};
