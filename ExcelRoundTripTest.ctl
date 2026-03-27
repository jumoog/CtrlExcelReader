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
  excelRoundTripTest();
}

// Minimal round-trip smoke test for CtrlExcelReader.
// Writes an .xlsx using excelWriteFile() and reads it back using excelReadFile().

bool excelRoundTripTest(string filename = "")
{
  // Pick a temporary file if none provided.
  if (filename == "")
  {
    filename = tmpnam();

    if (filename == "")
    {
      DebugTN("excelRoundTripTest: tmpnam failed");
      return FALSE;
    }
  }

  // Build sample data (one sheet: "People")
  dyn_anytype rows;
  mapping row;
  time now = getCurrentTime();
  // it will fail with milliseconds
  time t1  = makeTime(year(now), month(now), day(now), hour(now), minute(now), second(now));
  time t2 = makeTime(2026, 1, 1, 1, 1, 1);
  row["Name"] = "Alice";
  row["Age"] = 30;
  row["Score"] = 95.5;
  row["Active"] = TRUE;
  row["Time"] = t1;
  dynAppend(rows, row);

  mapping row2;
  row2["Name"] = "Bob";
  row2["Age"] = 25;
  row2["Score"] = 87.0;
  row2["Active"] = FALSE;
  row2["Time"] = t2;
  dynAppend(rows, row2);

  // Write
  bool ok =excelWriteSheet(filename, "People", rows);

  if (!ok)
  {
    DebugN("excelRoundTripTest: excelWriteFile failed", filename);
    return FALSE;
  }

  // Read back
  mapping back = excelReadFile(filename);

  bool pass = TRUE;

  // Value checks (headers enabled by writer)
  if (back["People"][1]["Name"] != "Alice") pass = FALSE;

  if (back["People"][1]["Age"] != 30) pass = FALSE;

  if (back["People"][1]["Active"] != TRUE) pass = FALSE;

  if (back["People"][1]["Time"] != t1) pass = FALSE;

  if (back["People"][2]["Name"] != "Bob") pass = FALSE;

  if (back["People"][2]["Age"] != 25) pass = FALSE;

  if (back["People"][2]["Active"] != FALSE) pass = FALSE;

  if (back["People"][2]["Time"] != t2) pass = FALSE;

  DebugTN("excelRoundTripTest", "file", filename, "pass", pass);

  if (!pass)
  {
    DebugTN("excelRoundTripTest: read-back data", back);
  }

  // Cleanup (best-effort)
  remove(filename);

  return pass;
}
