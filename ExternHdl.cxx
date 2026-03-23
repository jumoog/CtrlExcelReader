#include <ExternHdl.hxx>

#include <DynVar.hxx>
#include <FloatVar.hxx>
#include <MappingVar.hxx>

#include <xlsxio_read.h>

#include <cstdlib>
#include <climits>
#include <string>
#include <vector>

//------------------------------------------------------------------------------

// Set a mapping value using the cell type reported by xlsxio.
//   VALUE    -> IntegerVar or FloatVar (parsed from string)
//   BOOLEAN  -> BitVar
//   DATE     -> TextVar (kept as string)
//   STRING / NONE -> TextVar
static void setTypedCell(MappingVar &row, const char *key, xlsxioread_cell cell)
{
  const char *value = cell->data;

  if ( !value || !*value )
  {
    row.setAt(TextVar(key), TextVar(""));
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
        row.setAt(TextVar(key), IntegerVar(static_cast<int>(lval)));
        return;
      }
      double dval = strtod(value, &end);
      if ( *end == '\0' && end != value )
      {
        row.setAt(TextVar(key), FloatVar(dval));
        return;
      }
      row.setAt(TextVar(key), TextVar(value));
      return;
    }
    case XLSXIOREAD_CELL_TYPE_BOOLEAN:
    {
      row.setAt(TextVar(key), BitVar(value[0] != '0'));
      return;
    }

    default:
      row.setAt(TextVar(key), TextVar(value));
      return;
  }
}

// Read an open sheet into a DynVar of MappingVar rows.
// First row is treated as column headers used as mapping keys.
static void readSheetRows(xlsxioreadersheet sheet, DynVar &result)
{
  // Read the first row as column headers
  std::vector<std::string> headers;
  if ( xlsxioread_sheet_next_row(sheet) )
  {
    xlsxioread_cell cell;
    while ( (cell = xlsxioread_sheet_next_cell_struct(sheet)) != nullptr )
    {
      headers.emplace_back(cell->data ? cell->data : "");
      free(cell);
    }
  }

  // Read data rows, keyed by the header names
  while ( xlsxioread_sheet_next_row(sheet) )
  {
    MappingVar rowMap;
    int colIdx = 0;
    xlsxioread_cell cell;
    while ( (cell = xlsxioread_sheet_next_cell_struct(sheet)) != nullptr )
    {
      const char *key = (colIdx < (int)headers.size())
        ? headers[colIdx].c_str()
        : std::to_string(colIdx + 1).c_str();
      setTypedCell(rowMap, key, cell);
      free(cell);
      colIdx++;
    }
    result.append(rowMap);
  }
}

static FunctionListRec fnList[] =
{
  { DYNTEXT_VAR,       "excelGetSheetNames", "(string filename)",                                        false },
  { DYNMAPPING_VAR,    "excelReadSheet",     "(string filename, string sheetName, bool skipHiddenRows = true)", false },
  { DYNDYNMAPPING_VAR, "excelReadFile",      "(string filename, bool skipHiddenRows = true)",                   false },
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
  };

  static DynVar dynTextResult;
  static DynVar dynMappingResult;
  static DynVar dynDynMappingResult;

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

      readSheetRows(sheet, dynMappingResult);

      xlsxioread_sheet_close(sheet);
      xlsxioread_close(reader);

      return &dynMappingResult;
    }

    // -------------------------------------------------------------------------
    // excelReadFile(string filename, bool skipHiddenRows) -> dyn_dyn_mapping
    // Reads all sheets in the file. Each sheet becomes a dyn_mapping of rows.
    // skipHiddenRows is optional and defaults to TRUE.
    case F_excelReadFile:
    {
      param.thread->clearLastError();
      dynDynMappingResult.reset(DYNMAPPING_VAR);

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

      xlsxioreader reader = xlsxioread_open(filenameVar.getValue());
      if ( !reader )
        return &dynDynMappingResult;

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
        readSheetRows(sheet, sheetDyn);
        xlsxioread_sheet_close(sheet);

        dynDynMappingResult.append(sheetDyn);
      }

      xlsxioread_close(reader);
      return &dynDynMappingResult;
    }

    // -------------------------------------------------------------------------
    default:
      return &errorIntVar;
  }
}

//------------------------------------------------------------------------------
