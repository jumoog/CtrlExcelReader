#include <ExcelXlsxHelpers.hxx>

#include <BitVar.hxx>
#include <AnyTypeVar.hxx>
#include <DynVar.hxx>
#include <FloatVar.hxx>
#include <IntegerVar.hxx>
#include <LongVar.hxx>
#include <MappingVar.hxx>
#include <MixedVar.hxx>
#include <TextVar.hxx>
#include <TimeVar.hxx>

#include <cctype>
#include <climits>
#include <unordered_map>
#include <cmath>
#include <cstring>
#include <ctime>
#include <string>
#include <vector>

using namespace OpenXLSX;

namespace ExcelXlsxHelpers
{
  const Variable *unwrapAnyOrMixed(const Variable *val)
  {
    const Variable *current = val;

    while (current
        && (current->isA() == ANYTYPE_VAR || current->isA() == MIXED_VAR))
    {
      const AnyTypeVar *wrapped = static_cast<const AnyTypeVar *>(current);
      current = wrapped->getVar();
    }

    return current;
  }
}

namespace
{
  constexpr uint32_t EXCEL_FMT_DATE_TIME = 22;

  bool toLocalCalendarTime(time_t sec, std::tm &outTm)
  {
#ifdef _WIN32
    return localtime_s(&outTm, &sec) == 0;
#else
    return localtime_r(&sec, &outTm) != nullptr;
#endif
  }

  XLStyleIndex createBoldHeaderFormat(XLDocument &doc)
  {
    XLStyleIndex boldFontIdx = doc.styles().fonts().create(
      doc.styles().fonts().fontByIndex(0));
    doc.styles().fonts().fontByIndex(boldFontIdx).setBold(true);

    XLStyleIndex boldFmtIdx = doc.styles().cellFormats().create(
      doc.styles().cellFormats().cellFormatByIndex(0));
    doc.styles().cellFormats().cellFormatByIndex(boldFmtIdx)
      .setFontIndex(boldFontIdx);
    doc.styles().cellFormats().cellFormatByIndex(boldFmtIdx)
      .setApplyFont(true);

    return boldFmtIdx;
  }

  XLStyleIndex createDateTimeFormat(XLDocument &doc)
  {
    XLStyleIndex dateTimeFmtIdx = doc.styles().cellFormats().create(
      doc.styles().cellFormats().cellFormatByIndex(0));
    doc.styles().cellFormats().cellFormatByIndex(dateTimeFmtIdx)
      .setNumberFormatId(EXCEL_FMT_DATE_TIME);
    doc.styles().cellFormats().cellFormatByIndex(dateTimeFmtIdx)
      .setApplyNumberFormat(true);

    return dateTimeFmtIdx;
  }

  //----------------------------------------------------------------------------
  // Date-format detection helpers
  //   OpenXLSX reports dates as XLValueType::Float. We inspect the cell's
  //   number-format to tell dates from plain numbers.
  //----------------------------------------------------------------------------

  // Built-in Excel number-format IDs that represent dates/times (ECMA-376).
  bool isBuiltinDateFormatId(unsigned int id)
  {
    return (id >= 14 && id <= 22)
        || (id >= 27 && id <= 36)
        || (id >= 45 && id <= 47);
  }

  // Scan a custom format-code string for date/time tokens (y m d h s)
  // while ignoring quoted literals, escaped chars and bracketed sections.
  bool isDateFormatCode(const std::string &code)
  {
    bool inQuote = false;
    bool inBracket = false;

    for (size_t i = 0; i < code.size(); i++)
    {
      char c = code[i];

      if (c == '"') { inQuote = !inQuote; continue; }
      if (inQuote) continue;
      if (c == '\\') { i++; continue; } // skip escaped char
      if (c == '[') { inBracket = true; continue; }
      if (c == ']') { inBracket = false; continue; }
      if (inBracket) continue;

      char lower = static_cast<char>(tolower(static_cast<unsigned char>(c)));
      if (lower == 'y' || lower == 'm' || lower == 'd'
       || lower == 'h' || lower == 's')
        return true;
    }
    return false;
  }

  //----------------------------------------------------------------------------
  // Read helpers
  //----------------------------------------------------------------------------

  // Set a mapping value using the cell type reported by OpenXLSX, using a
  // pre-fetched XLStyles reference and a caller-owned format-index → is_date
  // cache to avoid redundant style lookups.
  void setTypedCellCached(MappingVar &row, const Variable &key,
                          XLCell &cell,
                          const XLStyles &styles,
                          std::unordered_map<XLStyleIndex, bool> &dateCache)
  {
    XLCellValue val = cell.value();
    auto type = val.type();

    switch (type)
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
        bool isDate = false;
        try
        {
          XLStyleIndex styleIdx = cell.cellFormat();
          auto it = dateCache.find(styleIdx);
          if (it != dateCache.end())
          {
            isDate = it->second;
          }
          else
          {
            auto fmt = styles.cellFormats().cellFormatByIndex(styleIdx);
            unsigned int fmtId = fmt.numberFormatId();
            isDate = isBuiltinDateFormatId(fmtId);
            if (!isDate && fmtId >= 164)
            {
              std::string code = styles.numberFormats()
                                       .numberFormatById(fmtId)
                                       .formatCode();
              isDate = isDateFormatCode(code);
            }
            dateCache[styleIdx] = isDate;
          }
        }
        catch (...) {}

        if (isDate)
        {
          XLDateTime dt = val.get<XLDateTime>();
          std::tm tm = dt.tm();
          tm.tm_isdst = -1;
          time_t sec = mktime(&tm);

          double serial = dt.serial();
          double frac = serial - floor(serial);

          // Convert fraction-of-day to milliseconds with rounding so
          // floating-point noise does not turn exact .000 into .999.
          const long long dayMillis = 24LL * 60LL * 60LL * 1000LL;
          long long totalMillis = static_cast<long long>(llround(frac * static_cast<double>(dayMillis)));

          if (totalMillis < 0)
            totalMillis = 0;
          else if (totalMillis >= dayMillis)
          {
            totalMillis = 0;
            sec += 1;
          }

          PVSSshort milli = static_cast<PVSSshort>(totalMillis % 1000LL);
          row.setAt(key, TimeVar(sec, milli));
        }
        else
        {
          double dval = val.get<double>();
          double intpart;
          if (modf(dval, &intpart) == 0.0
           && intpart >= INT_MIN && intpart <= INT_MAX)
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

  //----------------------------------------------------------------------------
  // Write helpers
  //----------------------------------------------------------------------------

  // Write a single WinCC OA Variable to an OpenXLSX cell.
  void writeTypedCell(XLCell &cell, const Variable *val,
                      XLStyleIndex dateTimeFmtIdx)
  {
    if (!val)
    {
      cell.value() = std::string();
      return;
    }

    switch (val->isA())
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

        // Excel date/time serials have no timezone. Convert epoch seconds to
        // local calendar fields first, then write those fields as XLDateTime
        // so displayed wall-clock time matches WinCC OA local time.
        std::tm localTm{};
        if (toLocalCalendarTime(sec, localTm))
          cell.value() = XLDateTime(localTm);
        else
          cell.value() = XLDateTime(sec);

        cell.setCellFormat(dateTimeFmtIdx);
        return;
      }
      case TEXT_VAR:
        cell.value() = std::string(
          static_cast<const TextVar *>(val)->getValue());
        return;
      case ANYTYPE_VAR:
      case MIXED_VAR:
      {
        const Variable *inner = ExcelXlsxHelpers::unwrapAnyOrMixed(val);
        if (!inner)
        {
          cell.value() = std::string();
          return;
        }

        writeTypedCell(cell, inner, dateTimeFmtIdx);
        return;
      }
      default:
      {
        CharString str = val->formatValue(CharString());
        cell.value() = std::string(str.c_str());
        return;
      }
    }
  }
} // namespace

namespace ExcelXlsxHelpers
{
  // Read an open worksheet into a DynVar of MappingVar rows.
  void readSheetRows(XLWorksheet &wks, XLDocument &doc,
                     DynVar &result, bool useHeaders, bool skipHidden)
  {
    uint32_t rowCount = wks.rowCount();
    uint16_t colCount = wks.columnCount();

    if (rowCount == 0 || colCount == 0)
      return;

    std::vector<std::string> headers;
    headers.reserve(colCount);
    uint32_t dataStartRow = 1;
  
    if (useHeaders)
    {
      for (auto& cell : wks.row(1).cells(colCount))
      {
        XLCellValue val = cell.value();
        if (val.type() == XLValueType::Empty)
          headers.emplace_back("");
        else
          headers.emplace_back(val.get<std::string>());
      }
      dataStartRow = 2;
    }

    // Pre-fetch styles once and cache format-index → is_date results so each
    // unique cell format is inspected only once across the entire sheet.
    XLStyles styles = doc.styles();
    std::unordered_map<XLStyleIndex, bool> dateCache;

    for (auto& xlRow : wks.rows(dataStartRow, rowCount))
    {
      if (skipHidden && xlRow.isHidden())
        continue;

      MappingVar rowMap;
      uint16_t c = 1;
      for (auto& cell : xlRow.cells(colCount))
      {
        if (useHeaders && (c - 1) < static_cast<uint16_t>(headers.size()))
          setTypedCellCached(rowMap, TextVar(headers[c - 1].c_str()), cell, styles, dateCache);
        else
          setTypedCellCached(rowMap, IntegerVar(c), cell, styles, dateCache);
        ++c;
      }
      result.append(rowMap);
    }
  }

  // Write a DynVar of MappingVars to an OpenXLSX worksheet.
  // Column headers are taken from the keys of the first mapping row.
  bool writeSheetData(XLWorksheet &wks, DynVar &data, XLDocument &doc)
  {
    unsigned int numRows = data.getNumberOfItems();
    if (numRows == 0)
      return true;

    Variable *firstRowVar = data.getAt(0);
    if (!firstRowVar)
      return false;

    MappingVar firstRow;
    firstRow = *firstRowVar;
    unsigned int numCols = firstRow.getNumberOfItems();
    if (numCols == 0)
      return true;

    // Collect column key names from the first row
    std::vector<CharString> columnNames;
    columnNames.reserve(numCols);
    for (unsigned int c = 0; c < numCols; c++)
    {
      Variable *key = firstRow.getKey(c);
      columnNames.push_back(key->formatValue(CharString()));
    }

    // Track max character width per column (init with header lengths)
    std::vector<size_t> maxWidths(numCols);
    for (unsigned int c = 0; c < numCols; c++)
      maxWidths[c] = strlen(columnNames[c].c_str());

    // Create bold cell format for the header row.
    // XLFont/XLCellFormat are XML-node proxies: call create() first to duplicate
    // the default entry, then modify the new entry so the default is not mutated.
    XLStyleIndex boldFmtIdx = createBoldHeaderFormat(doc);

    // Built-in format 22: date + time display. Without a date format, Excel
    // shows the serial value (e.g. 46105.837...).
    XLStyleIndex dateTimeFmtIdx = createDateTimeFormat(doc);

    // Write column headers in row 1 (bold)
    for (unsigned int c = 0; c < numCols; c++)
    {
      auto cell = wks.cell(1, static_cast<uint16_t>(c + 1));
      cell.value() = std::string(columnNames[c].c_str());
      cell.setCellFormat(boldFmtIdx);
    }

    // Write data rows starting from row 2
    for (unsigned int r = 0; r < numRows; r++)
    {
      Variable *rowVar = data.getAt(r);
      if (!rowVar)
        continue;

      MappingVar row;
      row = *rowVar; // AnyTypeVar unwrapping via operator=

      for (unsigned int c = 0; c < numCols; c++)
      {
        TextVar keyVar(columnNames[c].c_str());
        Variable *cellVal = row.getAt(keyVar);
        auto cell = wks.cell(
          static_cast<uint32_t>(r + 2),
          static_cast<uint16_t>(c + 1)
        );
        writeTypedCell(cell, cellVal, dateTimeFmtIdx);

        if (cellVal)
        {
          CharString str = cellVal->formatValue(CharString());
          size_t len = strlen(str.c_str());
          if (len > maxWidths[c])
            maxWidths[c] = len;
        }
      }
    }

    // Set column widths (character width + padding)
    for (unsigned int c = 0; c < numCols; c++)
      wks.column(static_cast<uint16_t>(c + 1))
          .setWidth(static_cast<float>(maxWidths[c]) + 2.0f);

    return true;
  }
}
