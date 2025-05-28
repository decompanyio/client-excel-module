import { downloadExcel, uploadExcel } from ".";

const sampleData = {
  Sheet1: [
    { 이름: "홍길동", 나이: 30 },
    { 이름: "김철수", 나이: 25 },
  ],
  Sheet2: [{ 이름: "Tom", 나이: 40 }],
};

const downloadBtn = document.getElementById("downloadBtn");
const fileTypeSelect = document.getElementById("fileType") as HTMLSelectElement;

// 다운로드 버튼 클릭 시 선택된 타입으로 다운로드
downloadBtn?.addEventListener("click", () => {
  const fileType = fileTypeSelect.value as "xlsx" | "xls" | "csv" | "txt";
  downloadExcel(sampleData, "demo-data", fileType);
});

// 업로드 처리
const uploadInput = document.getElementById("uploadInput") as HTMLInputElement;

uploadInput?.addEventListener("change", async () => {
  const file = uploadInput.files?.[0];
  if (!file) return;

  try {
    const sheet = await uploadExcel(file);

    console.log("📥 업로드 성공:", sheet);
  } catch (error) {
    console.error("❌ 업로드 실패:", error);
  }
});
