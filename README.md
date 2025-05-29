# client-excel-module

엑셀 파일을 **업로드(파싱)** 및 **다운로드(생성)** 할 수 있는 TypeScript 모듈입니다.  
여러 워크시트, 멀티 헤더, 셀 스타일 등 다양한 엑셀 기능을 간편하게 지원합니다.

---

## 설치

```bash
pnpm install
```

## 주요 기능

- 엑셀 파일 업로드(파싱) 지원
  - 여러 워크시트 동시 파싱
  - 멀티 헤더(여러 줄 헤더) 지원
  - 원하는 행을 헤더로 지정 가능
- 엑셀 파일 다운로드(생성) 지원
  - 여러 워크시트 동시 생성
  - 멀티 헤더(여러 줄 헤더) 지원
  - 헤더/바디 셀 스타일 커스텀 지원
  - 다양한 포맷(xlsx, xls, csv, txt) 지원

## 데모

```bash
pnpm dev
```

`http://localhost:5173/demo/index.html` 접속.

demo 폴더 가면 데모용 ts 파일 확인 가능.

## 사용법

### 엑셀 업로드

```bash
// 기본 사용법 (모든 시트 파싱)
const file = ...; // File 객체
uploadExcel(file).then((sheets) => {
  // sheets = { Sheet1: [...], Sheet2: [...] }
});

// 특정 시트만 파싱
uploadExcel(file, { sheetName: "Sheet1" }).then((sheets) => {
  // sheets = { Sheet1: [...] }
});

// 멀티 헤더(2줄) 사용 예시
uploadExcel(file, { headerRowIndex: 0, headerRowCount: 2 }).then((sheets) => {
  // 헤더 2줄을 병합하여 데이터 파싱
});
```

### 업로드 옵션

| 옵션명           | 타입                   | 설명                                     | 기본값    |
| ---------------- | ---------------------- | ---------------------------------------- | --------- |
| `sheetName`      | `string` \| `string[]` | 읽을 시트명 또는 시트명 배열             | 전체 시트 |
| `headerRowIndex` | `number`               | 헤더가 시작되는 행의 인덱스 (0부터 시작) | `0`       |
| `headerRowCount` | `number`               | 헤더 행 개수 (멀티 헤더 지원 가능)       | `1`       |

## 엑셀 다운로드

```bash
// 기본 사용법 (단일 시트)
downloadExcel(
  { Sheet1: [{ 이름: "홍길동", 나이: 30 }, { 이름: "김철수", 나이: 25 }] },
  "sample"
);

// 멀티 헤더, 스타일, 여러 시트 예시
downloadExcel(
  {
    Sheet1: [{ 이름: "홍길동", 나이: 30 }, { 이름: "김철수", 나이: 25 }],
    Sheet2: [{ Name: "Tom", Age: 40 }]
  },
  "multi-sheet",
  "xlsx",
  {
    headerStyle: { fill: { fgColor: { rgb: "FF0000" } } }, // 헤더 배경 빨강
    bodyStyle: { font: { sz: 12 } } // 바디 폰트 크기 12
  },
  {
    Sheet1: [["이름", "나이"], ["Name", "Age"]],
    Sheet2: [["Name", "Age"]]
  }
);
```

### 다운로드 옵션

| 옵션명            | 타입                                      | 설명                                    | 기본값   |
| ----------------- | ----------------------------------------- | --------------------------------------- | -------- |
| `sheetsData`      | `Record<string, T[]>`                     | 시트별 데이터 객체                      | **필수** |
| `filename`        | `string`                                  | 저장할 파일명 (확장자 제외)             | **필수** |
| `type`            | `"xlsx"` \| `"xls"` \| `"csv"` \| `"txt"` | 저장할 파일 형식                        | `"xlsx"` |
| `styleOptions`    | `{ headerStyle?, bodyStyle? }`            | 헤더/바디 셀 스타일 설정 (선택 항목)    | 선택     |
| `multiHeadersMap` | `Record<string, string[][]>`              | 시트별 멀티 헤더 설정 (2차원 배열 형태) | 선택     |

### 멀티 헤더/스타일 예시

```bash
downloadExcel(
  {
    Sheet1: [
      { 이름: "홍길동", 나이: 30, 직업: "개발자" },
      { 이름: "김철수", 나이: 25, 직업: "디자이너" }
    ]
  },
  "styled",
  "xlsx",
  {
    headerStyle: { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "4F81BD" } } },
    bodyStyle: { font: { sz: 14 } }
  },
  {
    Sheet1: [
      ["이름", "나이", "직업"],
      ["Name", "Age", "Job"]
    ]
  }
);
```

### 참고

- csv, txt 포멧은 스타일이 적용되지 않음.
- 멀티 헤더는 2차원 배열로 전달하며, 데이터에는 헤더를 포함하지 않습니다.

## License

This project is licensed under the [MIT License](./LICENSE).
