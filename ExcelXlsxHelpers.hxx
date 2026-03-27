#ifndef _EXCEL_XLSX_HELPERS_HXX_
#define _EXCEL_XLSX_HELPERS_HXX_

#include <AnyTypeVar.hxx>
#include <DynVar.hxx>
#include <MixedVar.hxx>

#include <OpenXLSX.hpp>

class Variable;

namespace ExcelXlsxHelpers
{
  const Variable *unwrapAnyOrMixed(const Variable *val);

  void readSheetRows(OpenXLSX::XLWorksheet &wks, OpenXLSX::XLDocument &doc,
                     DynVar &result, bool useHeaders, bool skipHidden);

  bool writeSheetData(OpenXLSX::XLWorksheet &wks, DynVar &data,
                      OpenXLSX::XLDocument &doc);
}

#endif
