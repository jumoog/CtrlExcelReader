#include <ExternHdl.hxx>

#include <DynVar.hxx>
#include <FloatVar.hxx>
#include <LongVar.hxx>
#include <MappingVar.hxx>
#include <TimeVar.hxx>

#include <xlsxio_read.h>
#include <xlsxio_write.h>

#include <cstdlib>
#include <climits>
#include <cmath>
#include <string>
#include <vector>

//------------------------------------------------------------------------------

// Set a mapping value using the cell type reported by xlsxio.
//   VALUE    -> IntegerVar or FloatVar (parsed from string)
//   BOOLEAN  -> BitVar
//   DATE     -> TimeVar (Excel serial converted to time_t)
//   STRING / NONE -> TextVar
static void setTypedCell(MappingVar &row, const Variable &key, xlsxioread_cell cell)
{
  const char *value = cell->data;

  if ( !value || !*value )
  {
    row.setAt(key, TextVar(""));
    return;
  }

  switch ( cell->cell_type )
  {
    case XLSXIOREAD_CELL_TYPE_VALUE:
    {
      char *end = nullptr;
      long lval = strtol(value, &end, 10);
      if ( *end == '\0' && end != value && lval >= INT_MIN && lval <= INT_MAX )
      {
        row.setAt(key, IntegerVar(static_cast<int>(lval)));
        return;
      }
      double dval = strtod(value, &end);
      if ( *end == '\0' && end != value )
      {
        row.setAt(key, FloatVar(dval));
        return;
      }
      row.setAt(key, TextVar(value));
      return;
    }
    case XLSXIOREAD_CELL_TYPE_BOOLEAN:
    {
      row.setAt(key, BitVar(value[0] != '0'));
      return;
    }
    case XLSXIOREAD_CELL_TYPE_DATE:
    {
      // Excel stores dates as a serial number (days since Dec 30, 1899).
      // 25569 is the serial for Jan 1, 1970 (Unix epoch) in Excel's system,
      // which includes the historical Feb 29, 1900 leap-year bug.
      // Excel dates carry no timezone — treat them as local time.
      char *end = nullptr;
      double serial = strtod(value, &end);
      if ( *end == '\0' && end != value )
      {
        double totalSeconds = (serial - 25569.0) * 86400.0;
        time_t naive = static_cast<time_t>(floor(totalSeconds));
        double frac = totalSeconds - floor(totalSeconds);

        // Decompose as UTC, then reinterpret as local time via mktime
        struct tm components = *gmtime(&naive);
        components.tm_isdst = -1;
        time_t sec = mktime(&components);

        PVSSshort milli = static_cast<PVSSshort>(frac * 1000.0);
        row.setAt(key, TimeVar(sec, milli));
        return;
      }
      row.setAt(key, TextVar(value));
      return;
    }
    default:
      row.setAt(key, TextVar(value));
      return;
  }
}

// Read an open sheet into a DynVar of MappingVar rows.
// When useHeaders is true, the first row supplies the mapping keys;
// otherwise every row uses 1-based column numbers as keys.
static void readSheetRows(xlsxioreadersheet sheet, DynVar &result, bool useHeaders)
{
  std::vector<std::string> headers;

  if ( useHeaders )
  {
    // Read the first row as column headers
    if ( xlsxioread_sheet_next_row(sheet) )
    {
      xlsxioread_cell cell;
      while ( (cell = xlsxioread_sheet_next_cell_struct(sheet)) != nullptr )
      {
        headers.emplace_back(cell->data ? cell->data : "");
        free(cell);
      }
    }
  }

  // Read data rows
  while ( xlsxioread_sheet_next_row(sheet) )
  {
    MappingVar rowMap;
    int colIdx = 0;
    xlsxioread_cell cell;
    while ( (cell = xlsxioread_sheet_next_cell_struct(sheet)) != nullptr )
    {
      if ( useHeaders && colIdx < (int)headers.size() )
        setTypedCell(rowMap, TextVar(headers[colIdx].c_str()), cell);
      else
        setTypedCell(rowMap, IntegerVar(colIdx + 1), cell);
      free(cell);
      colIdx++;
    }
    result.append(rowMap);
  }
}

//------------------------------------------------------------------------------

// Write a single cell value, choosing the xlsxio call by WinCC OA Variable type.
static void writeTypedCell(xlsxiowriter writer, const Variable *val)
{
  if ( !val )
  {
    xlsxiowrite_add_cell_string(writer, "");
    return;
  }

  switch ( val->isA() )
  {
    case INTEGER_VAR:
      xlsxiowrite_add_cell_int(writer, static_cast<int64_t>(
        static_cast<const IntegerVar *>(val)->getValue()));
      return;
    case LONG_VAR:
      xlsxiowrite_add_cell_int(writer, static_cast<int64_t>(
        static_cast<const LongVar *>(val)->getValue()));
      return;
    case FLOAT_VAR:
      xlsxiowrite_add_cell_float(writer,
        static_cast<const FloatVar *>(val)->getValue());
      return;
    case BIT_VAR:
      xlsxiowrite_add_cell_int(writer,
        static_cast<const BitVar *>(val)->isTrue() ? 1 : 0);
      return;
    case TIME_VAR:
      xlsxiowrite_add_cell_datetime(writer, static_cast<time_t>(
        static_cast<const TimeVar *>(val)->getSeconds()));
      return;
    case TEXT_VAR:
      xlsxiowrite_add_cell_string(writer,
        static_cast<const TextVar *>(val)->getValue());
      return;
    default:
    {
      CharString str = val->formatValue(CharString());
      xlsxiowrite_add_cell_string(writer, str.c_str());
      return;
    }
  }
}

// Write a dyn_mapping (array of row-mappings) to an already-opened xlsxio writer.
// Column headers are taken from the keys of the first mapping row.
static bool writeSheetData(xlsxiowriter writer, const DynVar &data)
{
  unsigned int numRows = data.getNumberOfItems();
  if ( numRows == 0 )
    return true;

  Variable *firstRowVar = data.getAt(0);
  if ( !firstRowVar || firstRowVar->isA() != MAPPING_VAR )
    return false;

  const MappingVar *firstRow = static_cast<const MappingVar *>(firstRowVar);
  unsigned int numCols = firstRow->getNumberOfItems();
  if ( numCols == 0 )
    return true;

  // Collect column keys from the first row
  std::vector<Variable *> columnKeys;
  columnKeys.reserve(numCols);
  for ( unsigned int c = 0; c < numCols; c++ )
    columnKeys.push_back(firstRow->getKey(c));

  // Write column headers
  for ( Variable *key : columnKeys )
  {
    CharString keyStr = key->formatValue(CharString());
    xlsxiowrite_add_column(writer, keyStr.c_str(), 0);
  }
  xlsxiowrite_next_row(writer);

  // Write data rows
  for ( unsigned int r = 0; r < numRows; r++ )
  {
    Variable *rowVar = data.getAt(r);
    if ( !rowVar || rowVar->isA() != MAPPING_VAR )
    {
      xlsxiowrite_next_row(writer);
      continue;
    }

    const MappingVar *row = static_cast<const MappingVar *>(rowVar);
    for ( Variable *key : columnKeys )
    {
      Variable *cellVal = row->getAt(*key);
      writeTypedCell(writer, cellVal);
    }
    xlsxiowrite_next_row(writer);
  }

  return true;
}

static FunctionListRec fnList[] =
{
  { DYNTEXT_VAR,       "excelGetSheetNames", "(string filename)",                                                                          false },
  { DYNMAPPING_VAR,    "excelReadSheet",     "(string filename, string sheetName, bool skipHiddenRows = true, bool firstRowIsColumnNames = true)", false },
  { MAPPING_VAR,       "excelReadFile",      "(string filename, bool skipHiddenRows = true, bool firstRowIsColumnNames = true)",                   false },
  { BIT_VAR,           "excelWriteSheet",    "(string filename, string sheetName, dyn_mapping data)",                                              false },
  { BIT_VAR,           "excelWriteFile",     "(string filename, mapping data)",                                                                    false },
};

CTRL_EXTENSION(ExternHdl, fnList)

//------------------------------------------------------------------------------

const Variable *ExternHdl::execute(ExecuteParamRec &param)
{
  enum
  {
    F_excelGetSheetNames = 0,
    F_excelReadSheet     = 1,
    F_excelReadFile      = 2,
    F_excelWriteSheet    = 3,
    F_excelWriteFile     = 4,
  };

  static DynVar dynTextResult;
  static DynVar dynMappingResult;
  static MappingVar mappingResult;
  static BitVar writeResult;

  switch ( param.funcNum )
  {
    // -------------------------------------------------------------------------
    // excelGetSheetNames(string filename) -> dyn_string
    // Returns the names of all sheets in the given .xlsx file.
    case F_excelGetSheetNames:
    {
      param.thread->clearLastError();
      dynTextResult.reset(TEXT_VAR);

      if ( !hasNumArgs(1, param) )
        return &dynTextResult;

      TextVar filenameVar;
      filenameVar = *(param.args->getFirst()->evaluate(param.thread));

      xlsxioreader reader = xlsxioread_open(filenameVar.getValue());
      if ( !reader )
        return &dynTextResult;

      xlsxioreadersheetlist sheetlist = xlsxioread_sheetlist_open(reader);
      const XLSXIOCHAR *sheetname;
      while ( (sheetname = xlsxioread_sheetlist_next(sheetlist)) != nullptr )
        dynTextResult.append(TextVar(sheetname));
      xlsxioread_sheetlist_close(sheetlist);
      xlsxioread_close(reader);

      return &dynTextResult;
    }

    // -------------------------------------------------------------------------
    // excelReadSheet(string filename, string sheetName, bool skipHiddenRows)
    //   -> dyn_mapping
    // Reads a sheet as a dyn_mapping. The first row is used as header;
    // each subsequent row becomes a mapping keyed by the header values.
    // Pass an empty sheetName to read the first sheet.
    // skipHiddenRows: when TRUE, hidden rows are omitted from the result.
    case F_excelReadSheet:
    {
      param.thread->clearLastError();
      dynMappingResult.reset(MAPPING_VAR);

      TextVar filenameVar, sheetnameVar;
      filenameVar  = *(param.args->getFirst()->evaluate(param.thread));
      sheetnameVar = *(param.args->getNext() ->evaluate(param.thread));

      bool skipHidden = true;
      CtrlExpr *skipArg = param.args->getNext();
      if ( skipArg )
      {
        BitVar skipHiddenVar;
        skipHiddenVar = *(skipArg->evaluate(param.thread));
        skipHidden = skipHiddenVar.isTrue();
      }

      bool useHeaders = true;
      CtrlExpr *headerArg = param.args->getNext();
      if ( headerArg )
      {
        BitVar headerVar;
        headerVar = *(headerArg->evaluate(param.thread));
        useHeaders = headerVar.isTrue();
      }

      xlsxioreader reader = xlsxioread_open(filenameVar.getValue());
      if ( !reader )
        return &dynMappingResult;

      unsigned flags = XLSXIOREAD_SKIP_EMPTY_ROWS;
      if ( skipHidden )
        flags |= XLSXIOREAD_SKIP_HIDDEN_ROWS;

      // Pass nullptr to open the first sheet when sheetName is empty
      const char *sheetname = sheetnameVar.getValue();
      xlsxioreadersheet sheet = xlsxioread_sheet_open(
        reader,
        (*sheetname ? sheetname : nullptr),
        flags
      );

      if ( !sheet )
      {
        xlsxioread_close(reader);
        return &dynMappingResult;
      }

      readSheetRows(sheet, dynMappingResult, useHeaders);

      xlsxioread_sheet_close(sheet);
      xlsxioread_close(reader);

      return &dynMappingResult;
    }

    // -------------------------------------------------------------------------
    // excelReadFile(string filename, bool skipHiddenRows, bool firstRowIsColumnNames)
    //   -> mapping
    // Reads all sheets. Returns a mapping where keys are sheet names
    // and values are dyn_mapping (rows).
    case F_excelReadFile:
    {
      param.thread->clearLastError();
      mappingResult = MappingVar();

      TextVar filenameVar;
      filenameVar = *(param.args->getFirst()->evaluate(param.thread));

      bool skipHidden = true;
      CtrlExpr *skipArg = param.args->getNext();
      if ( skipArg )
      {
        BitVar skipHiddenVar;
        skipHiddenVar = *(skipArg->evaluate(param.thread));
        skipHidden = skipHiddenVar.isTrue();
      }

      bool useHeaders = true;
      CtrlExpr *headerArg = param.args->getNext();
      if ( headerArg )
      {
        BitVar headerVar;
        headerVar = *(headerArg->evaluate(param.thread));
        useHeaders = headerVar.isTrue();
      }

      xlsxioreader reader = xlsxioread_open(filenameVar.getValue());
      if ( !reader )
        return &mappingResult;

      // Collect sheet names
      std::vector<std::string> sheetNames;
      xlsxioreadersheetlist sheetlist = xlsxioread_sheetlist_open(reader);
      const XLSXIOCHAR *name;
      while ( (name = xlsxioread_sheetlist_next(sheetlist)) != nullptr )
        sheetNames.emplace_back(name);
      xlsxioread_sheetlist_close(sheetlist);

      unsigned flags = XLSXIOREAD_SKIP_EMPTY_ROWS;
      if ( skipHidden )
        flags |= XLSXIOREAD_SKIP_HIDDEN_ROWS;

      for ( const auto &sn : sheetNames )
      {
        xlsxioreadersheet sheet = xlsxioread_sheet_open(reader, sn.c_str(), flags);
        if ( !sheet )
          continue;

        DynVar sheetDyn;
        sheetDyn.reset(MAPPING_VAR);
        readSheetRows(sheet, sheetDyn, useHeaders);
        xlsxioread_sheet_close(sheet);

        mappingResult.setAt(TextVar(sn.c_str()), sheetDyn);
      }

      xlsxioread_close(reader);
      return &mappingResult;
    }

    // -------------------------------------------------------------------------
    // excelWriteSheet(string filename, string sheetName, dyn_mapping data)
    //   -> bool
    // Writes a dyn_mapping to a single-sheet .xlsx file.
    // Mapping keys from the first row become column headers.
    case F_excelWriteSheet:
    {
      param.thread->clearLastError();
      writeResult = BitVar(false);

      TextVar filenameVar, sheetnameVar;
      filenameVar  = *(param.args->getFirst()->evaluate(param.thread));
      sheetnameVar = *(param.args->getNext() ->evaluate(param.thread));

      DynVar dataVar;
      dataVar = *(param.args->getNext()->evaluate(param.thread));

      const char *sheetname = sheetnameVar.getValue();
      xlsxiowriter writer = xlsxiowrite_open(
        filenameVar.getValue(),
        (*sheetname ? sheetname : "Sheet1")
      );
      if ( !writer )
        return &writeResult;

      bool ok = writeSheetData(writer, dataVar);
      int closeResult = xlsxiowrite_close(writer);

      if ( ok && closeResult == 0 )
        writeResult = BitVar(true);

      return &writeResult;
    }

    // -------------------------------------------------------------------------
    // excelWriteFile(string filename, mapping data) -> bool
    // Writes sheets from a mapping where keys are sheet names and
    // values are dyn_mapping (rows).
    // xlsxio only supports one sheet per file, so only the first entry is written.
    case F_excelWriteFile:
    {
      param.thread->clearLastError();
      writeResult = BitVar(false);

      TextVar filenameVar;
      filenameVar = *(param.args->getFirst()->evaluate(param.thread));

      MappingVar dataVar;
      dataVar = *(param.args->getNext()->evaluate(param.thread));

      unsigned int numSheets = dataVar.getNumberOfItems();
      if ( numSheets == 0 )
      {
        writeResult = BitVar(true);
        return &writeResult;
      }

      // Use the first entry: key = sheet name, value = dyn_mapping
      Variable *sheetNameKey = dataVar.getKey(0);
      Variable *sheetDataVar = dataVar.getValue(0);

      if ( !sheetNameKey || !sheetDataVar || !sheetDataVar->isDynVar() )
        return &writeResult;

      CharString sheetNameStr = sheetNameKey->formatValue(CharString());
      const DynVar *sheetData = static_cast<const DynVar *>(sheetDataVar);

      xlsxiowriter writer = xlsxiowrite_open(
        filenameVar.getValue(),
        sheetNameStr.c_str()
      );
      if ( !writer )
        return &writeResult;

      bool ok = writeSheetData(writer, *sheetData);
      int closeResult = xlsxiowrite_close(writer);

      if ( ok && closeResult == 0 )
        writeResult = BitVar(true);

      return &writeResult;
    }

    // -------------------------------------------------------------------------
    default:
      return &errorIntVar;
  }
}

//------------------------------------------------------------------------------
