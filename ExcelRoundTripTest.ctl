// $License: NOLICENSE
//--------------------------------------------------------------------------------
/**
  @file $relPath
  @copyright $copyright
  @author Kilian von Pflugk
*/

//--------------------------------------------------------------------------------
// Libraries used (#uses)
#uses "CtrlExcelReader"


//--------------------------------------------------------------------------------
// Variables and Constants

//--------------------------------------------------------------------------------
/**
*/
void main()
{
  excelRoundTripTestSingle();
  excelRoundTripTestFile();
}

// Returns a temporary filename, or "" on failure.
string getTempFile(string context)
{
  string filename = tmpnam();

  if (filename == "")
  {
    DebugTN(context + ": tmpnam failed");
  }

  return filename;
}

// Builds the shared two-row test dataset.
// t1/t2 are set to the time values written, so callers can compare on read-back.
dyn_anytype buildTestRows(time &t1, time &t2)
{
  time now = getCurrentTime();
  // Strip milliseconds — Excel serial dates have second resolution.
  t1 = makeTime(year(now), month(now), day(now), hour(now), minute(now), second(now));
  t2 = makeTime(2026, 1, 1, 1, 1, 1);

  mapping row1;
  row1["Name"]   = "Alice";
  row1["Age"]    = 30;
  row1["Score"]  = 95.5;
  row1["Active"] = TRUE;
  row1["Time"]   = t1;

  mapping row2;
  row2["Name"]   = "Bob";
  row2["Age"]    = 25;
  row2["Score"]  = 87.0;
  row2["Active"] = FALSE;
  row2["Time"]   = t2;

  dyn_anytype rows;
  dynAppend(rows, row1);
  dynAppend(rows, row2);
  return rows;
}

// Verifies read-back rows against the expected values written by buildTestRows().
bool checkRows(dyn_mapping rows, time t1, time t2, string context = "checkRows")
{
  if (dynlen(rows) != 2)
  {
    DebugTN(context + ": unexpected row count", dynlen(rows));
    return false;
  }

  dyn_string missingKeys;
  dyn_string keysToCheck = makeDynString("Name", "Age", "Active", "Time");

  for (int row = 1; row <= 2; row++)
  {
    for (int k = 1; k <= dynlen(keysToCheck); k++)
    {
      if (!mappingHasKey(rows[row], keysToCheck[k]))
      {
        dynAppend(missingKeys, "row" + row + "." + keysToCheck[k]);
      }
    }
  }

  if (dynlen(missingKeys) > 0)
  {
    DebugTN(context + ": missing keys", missingKeys);
    return false;
  }

  bool pass = rows[1]["Name"]   == "Alice"
              && rows[1]["Age"]    == 30
              && rows[1]["Active"] == TRUE
              && rows[1]["Time"]   == t1
              && rows[2]["Name"]   == "Bob"
              && rows[2]["Age"]    == 25
              && rows[2]["Active"] == FALSE
              && rows[2]["Time"]   == t2;

  if (!pass)
  {
    DebugTN(context + ": value mismatch — read-back data", rows, "expected t1", t1, "expected t2", t2);
  }

  return pass;
}

// Round-trip test for excelWriteSheet / excelGetSheetNames / excelReadSheet.
bool excelRoundTripTestSingle(string filename = "")
{
  if (filename == "")
  {
    filename = getTempFile("excelRoundTripTestSingle");

    if (filename == "") return FALSE;
  }

  time t1, t2;
  dyn_anytype rows = buildTestRows(t1, t2);

  if (!excelWriteSheet(filename, "People", rows))
  {
    DebugTN("excelRoundTripTestSingle: excelWriteSheet failed", filename);
    return FALSE;
  }

  if (!excelGetSheetNames(filename).contains("People"))
  {
    DebugTN("excelRoundTripTestSingle: excelGetSheetNames failed", filename);
    return FALSE;
  }

  bool pass = checkRows(excelReadSheet(filename, "People"), t1, t2, "excelRoundTripTestSingle");
  DebugTN("excelRoundTripTestSingle", "file", filename, "pass", pass);
  remove(filename);
  return pass;
}

// Round-trip test for excelWriteFile / excelReadFile (multi-sheet API).
bool excelRoundTripTestFile(string filename = "")
{
  if (filename == "")
  {
    filename = getTempFile("excelRoundTripTestFile");

    if (filename == "") return FALSE;
  }

  time t1, t2;
  mapping data;
  data["People"] = buildTestRows(t1, t2);

  if (!excelWriteFile(filename, data))
  {
    DebugTN("excelRoundTripTestFile: excelWriteFile failed", filename);
    return FALSE;
  }

  mapping back = excelReadFile(filename);

  if (!mappingHasKey(back, "People"))
  {
    DebugTN("excelRoundTripTestFile: sheet 'People' missing from read-back", mappingKeys(back));
    return false;
  }

  bool pass = checkRows(back["People"], t1, t2, "excelRoundTripTestFile");
  DebugTN("excelRoundTripTestFile", "file", filename, "pass", pass);
  remove(filename);
  return pass;
}
