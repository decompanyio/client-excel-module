# 📦 client-excel-module

> 엑셀 파일을 **업로드(파싱)** 및 **다운로드(생성)** 할 수 있는 TypeScript 클라이언트 모듈입니다.
> `xlsx`, `xlsx-js-style` 기반으로 여러 워크시트, 멀티 헤더, 셀 스타일 등 다양한 엑셀 기능을 간편하게 지원합니다.

---

## ✨ 주요 기능

- ✅ 엑셀 파일 업로드 (파싱)

  - 여러 워크시트 동시 파싱
  - 멀티 헤더(2줄 이상) 지원
  - 원하는 행을 헤더로 지정 가능

- ✅ 엑셀 파일 다운로드 (생성)

  - 여러 워크시트 생성
  - 멀티 헤더, 스타일 지정
  - 다양한 포맷 지원 (`xlsx`, `xls`, `csv`, `txt`)

---

## 📦 설치

```bash
pnpm install client-excel-module
# 또는
npm install client-excel-module
```

---

## 📤 엑셀 업로드 사용법

```ts
import { uploadExcel } from "client-excel-module";

// 기본 사용법 (모든 시트 파싱)
uploadExcel(file).then((sheets) => {
  console.log(sheets); // { Sheet1: [...], Sheet2: [...] }
});

// 특정 시트만 파싱
uploadExcel(file, { sheetName: "Sheet1" });

// 멀티 헤더 사용 (2줄 이상 헤더 병합)
uploadExcel(file, { headerRowIndex: 0, headerRowCount: 2 });
```

### 업로드 옵션

| 옵션명           | 타입                   | 설명                            | 기본값    |
| ---------------- | ---------------------- | ------------------------------- | --------- |
| `sheetName`      | `string` \| `string[]` | 읽을 시트명 또는 배열           | 전체 시트 |
| `headerRowIndex` | `number`               | 헤더가 시작되는 행 (0부터 시작) | `0`       |
| `headerRowCount` | `number`               | 헤더 행 수 (멀티 헤더 지원)     | `1`       |

---

## 📥 엑셀 다운로드 사용법

```ts
import { downloadExcel } from "client-excel-module";

// 기본 사용법 (단일 시트)
downloadExcel(
  {
    Sheet1: [
      { 이름: "홍길동", 나이: 30 },
      { 이름: "김철수", 나이: 25 },
    ],
  },
  "sample"
);

// 고급 사용법 (멀티 시트 + 멀티 헤더 + 스타일)
downloadExcel(
  {
    Sheet1: [{ 이름: "홍길동", 나이: 30 }],
    Sheet2: [{ Name: "Tom", Age: 40 }],
  },
  "multi-sheet",
  "xlsx",
  {
    headerStyle: {
      fill: { fgColor: { rgb: "FF0000" } },
    },
    bodyStyle: {
      font: { sz: 12 },
    },
  },
  {
    Sheet1: [
      ["이름", "나이"],
      ["Name", "Age"],
    ],
    Sheet2: [["Name", "Age"]],
  }
);
```

### 다운로드 옵션

| 옵션명            | 타입                                      | 설명                          | 기본값   |
| ----------------- | ----------------------------------------- | ----------------------------- | -------- |
| `sheetsData`      | `Record<string, T[]>`                     | 시트별 데이터 객체            | **필수** |
| `filename`        | `string`                                  | 저장할 파일명 (확장자 제외)   | **필수** |
| `type`            | `"xlsx"` \| `"xls"` \| `"csv"` \| `"txt"` | 저장할 파일 형식              | `"xlsx"` |
| `styleOptions`    | `{ headerStyle?, bodyStyle? }`            | 셀 스타일 지정 (선택)         | 선택     |
| `multiHeadersMap` | `Record<string, string[][]>`              | 시트별 멀티 헤더 (2차원 배열) | 선택     |

### 참고사항

- `csv`, `txt` 포맷은 스타일 적용 불가
- `multiHeadersMap`은 헤더 문자열만 전달 (실제 데이터는 헤더 없이)

---

## 🧪 데모

```bash
pnpm dev
```

브라우저에서 [http://localhost:5173/demo/index.html](http://localhost:5173/demo/index.html) 접속
`/demo` 폴더에서 데모용 TypeScript 파일 확인 가능

---

## 📄 License

MIT
