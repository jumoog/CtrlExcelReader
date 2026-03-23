# WinCC OA CTRL Extension for Excel

A WinCC OA CTRL extension that adds `.xlsx` file reading capabilities using a [fork of xlsxio](https://github.com/jumoog/xlsxio) with cell type detection.

## CTRL Functions

### `excelGetSheetNames`

```ctrl
dyn_string excelGetSheetNames(string filename)
```

Returns the names of all sheets in the given `.xlsx` file.

```ctrl
dyn_string sheets = excelGetSheetNames("C:/data/report.xlsx");
// sheets = {"Sheet1", "Sheet2", "Summary"}
```

### `excelReadSheet`

```ctrl
dyn_mapping excelReadSheet(string filename, string sheetName, bool skipHiddenRows)
```

Reads a sheet and returns each data row as a `mapping`. The first row is used as column headers (mapping keys). The `skipHiddenRows` parameter is optional and defaults to `TRUE`.

Pass an empty string for `sheetName` to read the first sheet.

Cell values are automatically typed based on the Excel cell type:

- Numeric integers → `int`
- Numeric decimals → `float`
- Booleans → `bool`
- Strings, dates, and everything else → `string`

```ctrl
// Read the first sheet, skip hidden rows (default)
dyn_mapping rows = excelReadSheet("C:/data/report.xlsx", "");

// Read a specific sheet, include hidden rows
dyn_mapping rows = excelReadSheet("C:/data/report.xlsx", "Sheet2", FALSE);

// Given an Excel sheet:
//   | Name  | Age | Score |
//   | Alice | 30  | 95.5  |
//   | Bob   | 25  | 87.0  |

DebugN(rows[1]["Name"]);   // "Alice"  (string)
DebugN(rows[1]["Age"]);    // 30       (int)
DebugN(rows[1]["Score"]);  // 95.5     (float)
```

### `excelReadFile`

```ctrl
dyn_dyn_mapping excelReadFile(string filename, bool skipHiddenRows)
```

Reads all sheets in the file at once. Returns a `dyn_dyn_mapping` where each element is one sheet (a `dyn_mapping` of rows). The `skipHiddenRows` parameter is optional and defaults to `TRUE`.

Use `excelGetSheetNames` to map sheet indices to names.

```ctrl
dyn_dyn_mapping allSheets = excelReadFile("C:/data/report.xlsx");

// allSheets[1] = first sheet rows, allSheets[2] = second sheet rows, ...
DebugN(allSheets[1][1]["Name"]);  // first row of first sheet
DebugN(allSheets[2][1]["Name"]);  // first row of second sheet
```

## Build

### Prerequisites

- WinCC OA 3.20 API
- [vcpkg](https://github.com/microsoft/vcpkg)
- CMake 3.15+
- Visual Studio (Windows)

### Install dependencies

```sh
vcpkg install --triplet x64-windows-static-md
```

### Configure and build

```sh
cmake -B build ^
  -DCMAKE_TOOLCHAIN_FILE=C:/Repos/vcpkg/scripts/buildsystems/vcpkg.cmake ^
  -DVCPKG_TARGET_TRIPLET=x64-windows-static-md

cmake --build build --config RelWithDebInfo
```

### Install

Copy the built `CtrlExcelReader.dll` from `build/RelWithDebInfo/` into your WinCC OA project's `bin/` directory.

Add the extension to your WinCC OA config file:

```text
[ctrl]
LoadCtrlLibs = "CtrlExcelReader"
```
