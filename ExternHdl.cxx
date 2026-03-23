#include <ExternHdl.hxx>

#include <DynVar.hxx>
#include <FloatVar.hxx>
#include <LongVar.hxx>
#include <MappingVar.hxx>
#include <TimeVar.hxx>

#include <OpenXLSX.hpp>

#include <cctype>
#include <climits>
#include <cmath>
#include <ctime>
#include <string>
#include <vector>

using namespace OpenXLSX;

//------------------------------------------------------------------------------
// Date-format detection helpers
//   OpenXLSX reports dates as XLValueType::Float.  We inspect the cell's
//   number-format to tell dates from plain numbers.
//------------------------------------------------------------------------------

// Built-in Excel number-format IDs that represent dates/times (ECMA-376).
static bool isBuiltinDateFormatId(unsigned int id)
{
  return (id >= 14 && id <= 22)
      || (id >= 27 && id <= 36)
      || (id >= 45 && id <= 47);
}

// Scan a custom format-code string for date/time tokens (y m d h s)
// while ignoring quoted literals, escaped chars and bracketed sections.
static bool isDateFormatCode(const std::string &code)
{
  bool inQuote   = false;
  bool inBracket = false;

  for (size_t i = 0; i < code.size(); i++)
  {
    char c = code[i];

    if ( c == '"' )          { inQuote = !inQuote; continue; }
    if ( inQuote )           continue;
    if ( c == '\\' )         { i++; continue; }         // skip escaped char
    if ( c == '[' )          { inBracket = true;  continue; }
    if ( c == ']' )          { inBracket = false; continue; }
    if ( inBracket )         continue;

    char lower = static_cast<char>(tolower(static_cast<unsigned char>(c)));
    if ( lower == 'y' || lower == 'm' || lower == 'd'
      || lower == 'h' || lower == 's' )
      return true;
  }
  return false;
}

// Check whether a cell's number format indicates a date.
static bool isCellDate(XLDocument &doc, const XLCell &cell)
{
  try
  {
    auto styles        = doc.styles();
    auto styleIdx      = cell.cellFormat();
    auto fmt           = styles.cellFormats().cellFormatByIndex(styleIdx);
    unsigned int fmtId = fmt.numberFormatId();

    if ( isBuiltinDateFormatId(fmtId) )
      return true;

    if ( fmtId >= 164 )
    {
      std::string code = styles.numberFormats()
                                .numberFormatById(fmtId)
                                .formatCode();
      return isDateFormatCode(code);
    }
  }
  catch (...) {}

  return false;
}

//------------------------------------------------------------------------------
// Read helpers
//------------------------------------------------------------------------------

// Set a mapping value using the cell type reported by OpenXLSX.
static void setTypedCell(MappingVar &row, const Variable &key,
                         XLCell &cell, XLDocument &doc)
{
  XLCellValue val = cell.value();
  auto type = val.type();

  switch ( type )
  {
    case XLValueType::Empty:
      row.setAt(key, TextVar(""));
      return;

    case XLValueType::Boolean:
      row.setAt(key, BitVar(val.get<bool>()));
      return;

    case XLValueType::Integer:
      row.setAt(key, IntegerVar(static_cast<int>(val.get<int64_t>())));
      return;

    case XLValueType::Float:
    {
      if ( isCellDate(doc, cell) )
      {
        XLDateTime dt = val.get<XLDateTime>();
        std::tm    tm = dt.tm();
        tm.tm_isdst   = -1;
        time_t sec    = mktime(&tm);

        double serial = dt.serial();
        double frac   = serial - floor(serial);
        PVSSshort milli = static_cast<PVSSshort>(frac * 86400.0 * 1000.0
                          - floor(frac * 86400.0) * 1000.0);
        row.setAt(key, TimeVar(sec, milli));
      }
      else
      {
        double dval = val.get<double>();
        double intpart;
        if ( modf(dval, &intpart) == 0.0
          && intpart >= INT_MIN && intpart <= INT_MAX )
        {
          row.setAt(key, IntegerVar(static_cast<int>(intpart)));
        }
        else
        {
          row.setAt(key, FloatVar(dval));
        }
      }
      return;
    }

    case XLValueType::String:
      row.setAt(key, TextVar(val.get<std::string>().c_str()));
      return;

    default:
      row.setAt(key, TextVar(""));
      return;
  }
}

// Read an open worksheet into a DynVar of MappingVar rows.
static void readSheetRows(XLWorksheet &wks, XLDocument &doc,
                          DynVar &result, bool useHeaders, bool skipHidden)
{
  uint32_t rowCount = wks.rowCount();
  uint16_t colCount = wks.columnCount();

  if ( rowCount == 0 || colCount == 0 )
    return;

  std::vector<std::string> headers;
  uint32_t dataStartRow = 1;

  if ( useHeaders )
  {
    for ( uint16_t c = 1; c <= colCount; c++ )
    {
      auto cell = wks.cell(1, c);
      XLCellValue val = cell.value();
      if ( val.type() == XLValueType::Empty )
        headers.emplace_back("");
      else
        headers.push_back(val.get<std::string>());
    }
    dataStartRow = 2;
  }

  for ( uint32_t r = dataStartRow; r <= rowCount; r++ )
  {
    auto xlRow = wks.row(r);
    if ( skipHidden && xlRow.isHidden() )
      continue;

    MappingVar rowMap;
    for ( uint16_t c = 1; c <= colCount; c++ )
    {
      auto cell = wks.cell(r, c);
      if ( useHeaders && (c - 1) < static_cast<uint16_t>(headers.size()) )
        setTypedCell(rowMap, TextVar(headers[c - 1].c_str()), cell, doc);
      else
        setTypedCell(rowMap, IntegerVar(c), cell, doc);
    }
    result.append(rowMap);
  }
}

//------------------------------------------------------------------------------
// Write helpers
//------------------------------------------------------------------------------

// Write a single WinCC OA Variable to an OpenXLSX cell.
static void writeTypedCell(XLCell &cell, const Variable *val)
{
  if ( !val )
  {
    cell.value() = std::string();
    return;
  }

  switch ( val->isA() )
  {
    case INTEGER_VAR:
      cell.value() = static_cast<int64_t>(
        static_cast<const IntegerVar *>(val)->getValue());
      return;
    case LONG_VAR:
      cell.value() = static_cast<int64_t>(
        static_cast<const LongVar *>(val)->getValue());
      return;
    case FLOAT_VAR:
      cell.value() = static_cast<const FloatVar *>(val)->getValue();
      return;
    case BIT_VAR:
      cell.value() = static_cast<const BitVar *>(val)->isTrue();
      return;
    case TIME_VAR:
    {
      time_t sec = static_cast<time_t>(
        static_cast<const TimeVar *>(val)->getSeconds());
      cell.value() = XLDateTime(sec);
      return;
    }
    case TEXT_VAR:
      cell.value() = std::string(
        static_cast<const TextVar *>(val)->getValue());
      return;
    default:
    {
      CharString str = val->formatValue(CharString());
      cell.value() = std::string(str.c_str());
      return;
    }
  }
}

// Write a DynVar of MappingVars to an OpenXLSX worksheet.
// Column headers are taken from the keys of the first mapping row.
static bool writeSheetData(XLWorksheet &wks, DynVar &data)
{
  unsigned int numRows = data.getNumberOfItems();
  if ( numRows == 0 )
    return true;

  Variable *firstRowVar = data.getAt(0);
  if ( !firstRowVar )
    return false;

  MappingVar firstRow;
  firstRow = *firstRowVar;
  unsigned int numCols = firstRow.getNumberOfItems();
  if ( numCols == 0 )
    return true;

  // Collect column key names from the first row
  std::vector<CharString> columnNames;
  columnNames.reserve(numCols);
  for ( unsigned int c = 0; c < numCols; c++ )
  {
    Variable *key = firstRow.getKey(c);
    columnNames.push_back(key->formatValue(CharString()));
  }

  // Write column headers in row 1
  for ( unsigned int c = 0; c < numCols; c++ )
    wks.cell(1, static_cast<uint16_t>(c + 1)).value() =
      std::string(columnNames[c].c_str());

  // Write data rows starting from row 2
  for ( unsigned int r = 0; r < numRows; r++ )
  {
    Variable *rowVar = data.getAt(r);
    if ( !rowVar )
      continue;

    MappingVar row;
    row = *rowVar;  // AnyTypeVar unwrapping via operator=

    for ( unsigned int c = 0; c < numCols; c++ )
    {
      TextVar keyVar(columnNames[c].c_str());
      Variable *cellVal = row.getAt(keyVar);
      auto cell = wks.cell(
        static_cast<uint32_t>(r + 2),
        static_cast<uint16_t>(c + 1)
      );
      writeTypedCell(cell, cellVal);
    }
  }

  return true;
}

//------------------------------------------------------------------------------

static FunctionListRec fnList[] =
{
  { DYNTEXT_VAR,       "excelGetSheetNames", "(string filename)",                                                                          false },
  { DYNMAPPING_VAR,    "excelReadSheet",     "(string filename, string sheetName, bool skipHiddenRows = true, bool firstRowIsColumnNames = true)", false },
  { MAPPING_VAR,       "excelReadFile",      "(string filename, bool skipHiddenRows = true, bool firstRowIsColumnNames = true)",                   false },
  { BIT_VAR,           "excelWriteSheet",    "(string filename, string sheetName, dyn_anytype data)",                                              false },
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
    case F_excelGetSheetNames:
    {
      param.thread->clearLastError();
      dynTextResult.reset(TEXT_VAR);

      if ( !hasNumArgs(1, param) )
        return &dynTextResult;

      TextVar filenameVar;
      filenameVar = *(param.args->getFirst()->evaluate(param.thread));

      try
      {
        XLDocument doc;
        doc.open(filenameVar.getValue());
        auto names = doc.workbook().sheetNames();
        for ( const auto &name : names )
          dynTextResult.append(TextVar(name.c_str()));
        doc.close();
      }
      catch (...) {}

      return &dynTextResult;
    }

    // -------------------------------------------------------------------------
    // excelReadSheet(string filename, string sheetName,
    //                bool skipHiddenRows, bool firstRowIsColumnNames)
    //   -> dyn_mapping
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

      try
      {
        XLDocument doc;
        doc.open(filenameVar.getValue());

        const char *sheetname = sheetnameVar.getValue();
        auto wks = (*sheetname)
          ? doc.workbook().worksheet(std::string(sheetname))
          : doc.workbook().worksheet(1);

        readSheetRows(wks, doc, dynMappingResult, useHeaders, skipHidden);
        doc.close();
      }
      catch (...) {}

      return &dynMappingResult;
    }

    // -------------------------------------------------------------------------
    // excelReadFile(string filename, bool skipHiddenRows,
    //              bool firstRowIsColumnNames) -> mapping
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

      try
      {
        XLDocument doc;
        doc.open(filenameVar.getValue());

        auto sheetNames = doc.workbook().worksheetNames();
        for ( const auto &sn : sheetNames )
        {
          auto wks = doc.workbook().worksheet(sn);

          DynVar sheetDyn;
          sheetDyn.reset(MAPPING_VAR);
          readSheetRows(wks, doc, sheetDyn, useHeaders, skipHidden);

          mappingResult.setAt(TextVar(sn.c_str()), sheetDyn);
        }

        doc.close();
      }
      catch (...) {}

      return &mappingResult;
    }

    // -------------------------------------------------------------------------
    // excelWriteSheet(string filename, string sheetName, dyn_anytype data)
    //   -> bool
    case F_excelWriteSheet:
    {
      param.thread->clearLastError();
      writeResult = BitVar(false);

      TextVar filenameVar, sheetnameVar;
      filenameVar  = *(param.args->getFirst()->evaluate(param.thread));
      sheetnameVar = *(param.args->getNext() ->evaluate(param.thread));

      const Variable *dataPtr = param.args->getNext()->evaluate(param.thread);
      if ( !dataPtr || !dataPtr->isDynVar() )
        return &writeResult;

      DynVar *dataVar = const_cast<DynVar *>(static_cast<const DynVar *>(dataPtr));

      try
      {
        XLDocument doc;
        doc.create(filenameVar.getValue(), XLForceOverwrite);
        doc.setProperty(XLProperty::Creator, "WinCC OA");
        doc.setProperty(XLProperty::LastModifiedBy, "WinCC OA");

        const char *sheetname = sheetnameVar.getValue();
        std::string sheetName = (*sheetname) ? sheetname : "Sheet1";

        auto &wb = doc.workbook();
        wb.worksheet(1).setName(sheetName);

        auto wks = wb.worksheet(sheetName);
        bool ok = writeSheetData(wks, *dataVar);

        doc.save();
        doc.close();

        if ( ok )
          writeResult = BitVar(true);
      }
      catch (...) {}

      return &writeResult;
    }

    // -------------------------------------------------------------------------
    // excelWriteFile(string filename, mapping data) -> bool
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

      try
      {
        XLDocument doc;
        doc.create(filenameVar.getValue(), XLForceOverwrite);
        doc.setProperty(XLProperty::Creator, "WinCC OA");
        doc.setProperty(XLProperty::LastModifiedBy, "WinCC OA");

        auto &wb = doc.workbook();
        bool ok = true;

        for ( unsigned int s = 0; s < numSheets; s++ )
        {
          CharString sheetName = dataVar.getKey(s)->formatValue(CharString());
          std::string sheetNameStr(sheetName.c_str());

          if ( s == 0 )
            wb.worksheet(1).setName(sheetNameStr);
          else
            wb.addWorksheet(sheetNameStr);

          auto wks = wb.worksheet(sheetNameStr);

          Variable *sheetDataVar = dataVar.getValue(s);
          if ( !sheetDataVar || !sheetDataVar->isDynVar() )
            continue;

          DynVar *sheetData = const_cast<DynVar *>(
            static_cast<const DynVar *>(sheetDataVar));
          ok = writeSheetData(wks, *sheetData) && ok;
        }

        doc.save();
        doc.close();

        if ( ok )
          writeResult = BitVar(true);
      }
      catch (...) {}

      return &writeResult;
    }

    // -------------------------------------------------------------------------
    default:
      return &errorIntVar;
  }
}

//------------------------------------------------------------------------------
