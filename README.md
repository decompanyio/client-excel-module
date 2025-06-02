# client-excel-module

> A TypeScript client module for **uploading (parsing)** and **downloading (generating)** Excel files.
> Built on `xlsx` and `xlsx-js-style`, it supports multiple sheets, multi-row headers, cell styles, and more with ease.

---

## ✨ Key Features

- ✅ Excel File Upload (Parsing)

  - Parse multiple sheets simultaneously
  - Supports multi-row headers (2 or more rows)
  - Specify which row to use as the header

- ✅ Excel File Download (Generation)

  - Generate multiple worksheets
  - Supports multi-row headers and cell styling
  - Supports various formats (`xlsx`, `xls`, `csv`, `txt`)

---

## 📦 Installation

```bash
pnpm install client-excel-module
# or
npm install client-excel-module
```

---

## 📤 Uploading Excel Files

```ts
import { uploadExcel } from "client-excel-module";

// Basic usage (parse all sheets)
uploadExcel(file).then((sheets) => {
  console.log(sheets); // { Sheet1: [...], Sheet2: [...] }
});

// Parse a specific sheet
uploadExcel(file, { sheetName: "Sheet1" });

// Use multi-row headers (merge 2 or more header rows)
uploadExcel(file, { headerRowIndex: 0, headerRowCount: 2 });
```

### Upload Options

| Option           | Type                   | Description                                  | Default    |
| ---------------- | ---------------------- | -------------------------------------------- | ---------- |
| `sheetName`      | `string` \| `string[]` | Name(s) of the sheets to parse               | All sheets |
| `headerRowIndex` | `number`               | Row index where the header starts (0-based)  | `0`        |
| `headerRowCount` | `number`               | Number of header rows (for multi-row header) | `1`        |

---

## 📥 Downloading Excel Files

```ts
import { downloadExcel } from "client-excel-module";

// Basic usage (single sheet)
downloadExcel(
  {
    Sheet1: [
      { Name: "Hong Gil-dong", Age: 30 },
      { Name: "Kim Cheol-soo", Age: 25 },
    ],
  },
  "sample"
);

// Advanced usage (multiple sheets, multi-row headers, and styles)
downloadExcel(
  {
    Sheet1: [{ Name: "Hong Gil-dong", Age: 30 }],
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

### Download Options

| Option            | Type                                      | Description                            | Default      |
| ----------------- | ----------------------------------------- | -------------------------------------- | ------------ |
| `sheetsData`      | `Record<string, T[]>`                     | Object containing sheet data           | **Required** |
| `filename`        | `string`                                  | Filename to save (without extension)   | **Required** |
| `type`            | `"xlsx"` \| `"xls"` \| `"csv"` \| `"txt"` | File format to save                    | `"xlsx"`     |
| `styleOptions`    | `{ headerStyle?, bodyStyle? }`            | Optional cell styling options          | Optional     |
| `multiHeadersMap` | `Record<string, string[][]>`              | Multi-row headers per sheet (2D array) | Optional     |

### Notes

- Styles are **not supported** for `csv` and `txt` formats.
- `multiHeadersMap` should only include header strings (actual data must not include headers).

---

## 🧪 Demo

```bash
pnpm dev
```

Open your browser and navigate to:
[http://localhost:5173/demo/index.html](http://localhost:5173/demo/index.html)
You can check out the demo TypeScript files in the `/demo` folder.

---

## 📄 License

MIT

---
