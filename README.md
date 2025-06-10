# client-excel-module

> A TypeScript client module for **uploading (parsing)** and **downloading (generating)** Excel files.
> Built as an **ESM (ECMAScript Module)** package.
> Uses `XLSX` to more easily provide the various features you need.

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
  - Custom column headers, order, and per-column formatting with `sheetsCol`
  - Per-column and global cell styling using [xlsx-js-style](https://github.com/SheetJS/sheetjs-style#cell-styles) format
  - Flexible: If `sheetsCol` is omitted, columns are auto-generated from your data

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
downloadExcel({
  sheetsData: {
    Sheet1: [
      { name: "Hong Gil-dong", age: 30 },
      { name: "Kim Cheol-soo", age: 25 },
    ],
  },
  filename: "sample",
});

// Advanced usage (multiple sheets, multi-row headers, styles, and column formatting)
downloadExcel({
  sheetsData: {
    Sheet1: [
      { name: "Hong Gil-dong", age: 30 },
      { name: "Kim Cheol-soo", age: 25 },
    ],
    Sheet2: [{ name: "Tom", age: 40 }],
  },
  sheetsCol: {
    Sheet1: [
      {
        key: "name",
        header: "Name",
        formatter: (v) => v.toUpperCase(),
        style: { font: { color: { rgb: "FF0000" }, bold: true } }, // red, bold
      },
      {
        key: "age",
        header: "Age",
        formatter: (v) => `${v} years`,
        style: { alignment: { horizontal: "center" } }, // center align
      },
    ],
    Sheet2: [
      { key: "name", header: "Name" },
      { key: "age", header: "Age" },
    ],
  },
  filename: "multi-sheet",
  type: "xlsx",
  styleOptions: {
    headerStyle: {
      fill: { fgColor: { rgb: "FF0000" } },
    },
    bodyStyle: {
      font: { sz: 12 },
    },
  },
  multiHeadersMap: {
    Sheet1: [
      ["이름", "나이"],
      ["Name", "Age"],
    ],
    Sheet2: [["Name", "Age"]],
  },
});
```

### Download Options

| Option            | Type                                      | Description                                        | Default      |
| ----------------- | ----------------------------------------- | -------------------------------------------------- | ------------ |
| `sheetsData`      | `Record<string, T[]>`                     | Object containing sheet data                       | **Required** |
| `sheetsCol`       | `Record<string, ExcelColumn<T>[]>`        | Column definitions per sheet (optional, see below) | Optional     |
| `filename`        | `string`                                  | Filename to save (without extension)               | **Required** |
| `type`            | `"xlsx"` \| `"xls"` \| `"csv"` \| `"txt"` | File format to save                                | `"xlsx"`     |
| `styleOptions`    | `{ headerStyle?, bodyStyle? }`            | Optional cell styling options                      | Optional     |
| `multiHeadersMap` | `Record<string, string[][]>`              | Multi-row headers per sheet (2D array)             | Optional     |

---

### Column Formatting & Styling (`sheetsCol` and `styleOptions`)

You can customize how your Excel columns are exported using the `sheetsCol` option, and control global styles with `styleOptions`.  
All style objects follow the [xlsx-js-style](https://github.com/SheetJS/sheetjs-style#cell-styles) format.

#### What does `sheetsCol` provide?

- **Custom column headers** (display names)
- **Column order** control
- **Per-column value formatting** (e.g., uppercase, add units, etc.)
- **Per-column header styling** (e.g., color, alignment)
- If omitted, columns are generated automatically from your data object keys.

#### What does `styleOptions` provide?

- **Global header style** (`headerStyle`): applies to all header cells in all sheets.
- **Global body style** (`bodyStyle`): applies to all data cells in all sheets.

#### Style Priority

1. **Per-column style (`sheetsCol[].style`)**: Highest priority, applied to the last header row of each column.
2. **Global header style (`styleOptions.headerStyle`)**: Used for all header cells unless overridden by per-column style.
3. **Global body style (`styleOptions.bodyStyle`)**: Used for all data cells.

> All style objects use the [xlsx-js-style](https://github.com/SheetJS/sheetjs-style#cell-styles) format.

#### Example

```ts
downloadExcel({
  sheetsData: {
    Sheet1: [
      { name: "Hong Gil-dong", age: 30 },
      { name: "Kim Cheol-soo", age: 25 },
    ],
  },
  sheetsCol: {
    Sheet1: [
      {
        key: "name",
        header: "Name",
        formatter: (v) => v.toUpperCase(),
        style: { font: { color: { rgb: "FF0000" }, bold: true } },
        wch: 10,
      },
      {
        key: "age",
        header: "Age",
        formatter: (v) => `${v} years`,
        style: { alignment: { horizontal: "center" } },
      },
    ],
  },
  filename: "styled-sample",
  type: "xlsx",
  styleOptions: {
    headerStyle: { fill: { fgColor: { rgb: "4F81BD" } } },
    bodyStyle: { font: { sz: 12 } },
  },
});
```

#### `sheetsCol` Type

```ts
type ExcelColumn<T> = {
  key: keyof T; // The property name in your data
  header: string; // The column header to display in Excel
  formatter?: (value: any, row: T) => any; // Optional value formatter
  style?: ExcelCellStyle; // Optional per-column cell style (xlsx-js-style format)
  wch?: number; // Optional column wch (wch, character units)
};
```

#### sheetsCol Explanation

- `key`: The property name in your data object (required)
- `header`: The column header to display in Excel (required)
- `formatter`: Function to transform the cell value (optional)
- `style`: Per-column cell style (optional, [xlsx-js-style](https://github.com/SheetJS/sheetjs-style#cell-styles) format)
- `wch`: Column wch (optional, in characters/wch, e.g., 30 means about 30 characters wide)

#### Column wch Example

```ts
sheetsCol: {
  Sheet1: [
    {
      key: "name",
      header: "Name",
      wch: 30,
      style: { font: { color: { rgb: "FF0000" }, bold: true } },
    },
    {
      key: "age",
      header: "Age",
      wch: 10,
      style: { alignment: { horizontal: "center" } },
    },
  ],
}
```

**Note:**

- If `wch` is not specified, the default is 20 characters.
- The `style` property is only applied for `xlsx` and `xls` formats. For `csv` and `txt`, styles are ignored.
- If you omit `sheetsCol`, all properties in your data will be exported in their default order and with their original names as headers.

---

## Demo

```bash
pnpm dev
```

After running the above command and building the project, open your browser and go to:  
[http://localhost:5173/demo/basic/index.html](http://localhost:5173/demo/basic/index.html)

**Note:**  
The demo page works with the built files in the `dist` directory.  
Please make sure to build the project first (`pnpm build` or `npm run build`) before using the demo.

You can create your own demo or test pages by adding files under the `/demo` folder.

---

## License

MIT

---
