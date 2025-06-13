import * as XLSX from "xlsx-js-style";

/**
 * CSV/TXT 파일을 다운로드합니다.
 */
export const downloadTextFile = (
  workbook: XLSX.WorkBook,
  filename: string,
  type: "csv" | "txt"
): void => {
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
};

/**
 * Excel 파일을 다운로드합니다.
 */
export const downloadExcelFile = (
  workbook: XLSX.WorkBook,
  filename: string,
  type: "xlsx" | "xls"
): void => {
  XLSX.writeFile(workbook, `${filename}.${type}`);
};
