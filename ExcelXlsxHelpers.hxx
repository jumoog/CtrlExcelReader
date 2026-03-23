#ifndef _EXCEL_XLSX_HELPERS_HXX_
#define _EXCEL_XLSX_HELPERS_HXX_

#include <DynVar.hxx>

#include <OpenXLSX.hpp>

namespace ExcelXlsxHelpers
{
  void readSheetRows(OpenXLSX::XLWorksheet &wks, OpenXLSX::XLDocument &doc,
                     DynVar &result, bool useHeaders, bool skipHidden);

  bool writeSheetData(OpenXLSX::XLWorksheet &wks, DynVar &data,
                      OpenXLSX::XLDocument &doc);
}

#endif
