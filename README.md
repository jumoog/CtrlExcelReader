# WinCC OA CTRL Extension for Excel

A WinCC OA CTRL extension that adds `.xlsx` file reading and writing capabilities using a [fork of xlsxio](https://github.com/jumoog/xlsxio) with cell type detection.

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
dyn_mapping excelReadSheet(string filename, string sheetName, bool skipHiddenRows = TRUE, bool firstRowIsColumnNames = TRUE)
```

Reads a sheet and returns each data row as a `mapping`.

- **sheetName** — pass an empty string to read the first sheet.
- **skipHiddenRows** — optional, defaults to `TRUE`. When `TRUE`, hidden rows are omitted.
- **firstRowIsColumnNames** — optional, defaults to `TRUE`. When `TRUE`, the first row supplies the mapping keys as strings. When `FALSE`, keys are 1-based column integers.

Cell values are automatically typed based on the Excel cell type:

- Numeric integers → `int`
- Numeric decimals → `float`
- Booleans → `bool`
- Dates/times → `time` (Excel serial converted to WinCC OA time, treated as local time)
- Strings and everything else → `string`

```ctrl
// Read the first sheet with default options
dyn_mapping rows = excelReadSheet("C:/data/report.xlsx", "");

// Read a specific sheet, include hidden rows
dyn_mapping rows = excelReadSheet("C:/data/report.xlsx", "Sheet2", FALSE);

// Given an Excel sheet:
//   | Name  | Age | Score |
//   | Alice | 30  | 95.5  |
//   | Bob   | 25  | 87.0  |

// With firstRowIsColumnNames = TRUE (default):
DebugN(rows[1]["Name"]);   // "Alice"  (string)
DebugN(rows[1]["Age"]);    // 30       (int)
DebugN(rows[1]["Score"]);  // 95.5     (float)

// With firstRowIsColumnNames = FALSE:
dyn_mapping rows = excelReadSheet("C:/data/report.xlsx", "", TRUE, FALSE);
DebugN(rows[1][1]);  // "Name"   (string)
DebugN(rows[1][2]);  // "Age"    (string)
DebugN(rows[2][1]);  // "Alice"  (string)
DebugN(rows[2][2]);  // 30       (int)
```

### `excelReadFile`

```ctrl
mapping excelReadFile(string filename, bool skipHiddenRows = TRUE, bool firstRowIsColumnNames = TRUE)
```

Reads all sheets in the file at once. Returns a `mapping` where keys are sheet names and values are `dyn_mapping` (rows).

- **skipHiddenRows** — optional, defaults to `TRUE`.
- **firstRowIsColumnNames** — optional, defaults to `TRUE`.

```ctrl
mapping allSheets = excelReadFile("C:/data/report.xlsx");

// Access sheets by name
DebugN(allSheets["Sheet1"][1]["Name"]);   // first row of Sheet1
DebugN(allSheets["Summary"][1]["Total"]); // first row of Summary sheet
```

### `excelWriteSheet`

```ctrl
bool excelWriteSheet(string filename, string sheetName, dyn_anytype data)
```

Writes a `dyn_anytype` (containing mappings) to a single-sheet `.xlsx` file. Mapping keys from the first row become column headers.

```ctrl
dyn_mapping rows;
mapping row1, row2;
row1["Name"] = "Alice";
row1["Age"]  = 30;
row2["Name"] = "Bob";
row2["Age"]  = 25;
dynAppend(rows, row1);
dynAppend(rows, row2);

bool ok = excelWriteSheet("C:/data/output.xlsx", "People", rows);
```

### `excelWriteFile`

```ctrl
bool excelWriteFile(string filename, mapping data)
```

Writes a `mapping` where keys are sheet names and values are `dyn_anytype` (rows of mappings) to a multi-sheet `.xlsx` file. This is the same format returned by `excelReadFile`.

```ctrl
// Read and write back (round-trip)
mapping allSheets = excelReadFile("C:/data/input.xlsx");
bool ok = excelWriteFile("C:/data/output.xlsx", allSheets);

// Build from scratch
dyn_mapping rows;
mapping row;
row["Name"] = "Alice";
dynAppend(rows, row);

mapping data;
data["Sheet1"] = rows;
bool ok = excelWriteFile("C:/data/output.xlsx", data);
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
