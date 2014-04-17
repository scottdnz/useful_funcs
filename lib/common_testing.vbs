'common_testing.vbs

Sub AssertEqual(given, expected, procName)
  if not given = expected then
    Err.Raise vbObjectError + 99999, , procName & " failed"
  end if
End Sub

Sub AssertNotEqual(given, expected, procName)
  if not given <> expected then
    Err.Raise vbObjectError + 99999, , procName & " failed"
  end if
End Sub

Sub AssertGreater(given, expected, procName)
  if not given > expected then
    Err.Raise vbObjectError + 99999, , procName & " failed"
  end if
End Sub

Sub runReportTests(testProcs)
  cntr = 0
  For each testProc in testProcs
    Execute testProc
    wscript.echo "."
    cntr = cntr + 1
  Next
  wscript.echo cntr & " tests successfully passed."
End Sub

Function inArray(val, arr)
  Dim i
  Dim found
  found = False
  for i = 0 to Ubound(arr)
    If arr(i) = val Then
      found = True
      exit for
    End If
  next
  inArray = found
End Function

'Reads a CSV file and stores it in an array. The actual type returned is a 
'variant type
Function readCSVIntoArray(fPath, csvFile)
  Dim inCsvSys, inCsv 'Object(s)
  Dim rows() ' rowItems() 'Array(s)
  Dim rowCount 'Int(s)
  rowCount = -1
  'Read CSV file into array
  Set inCsvSys = CreateObject("Scripting.FileSystemObject") 
  Set inCsv = inCsvSys.OpenTextFile(fPath & csvFile, "1", True)
  inCsv.ReadLine 'Skip header row
  Do While Not inCsv.AtEndOfStream
    rowCount = rowCount + 1
    Redim Preserve rows(rowCount)
    rows(rowCount) = inCsv.ReadLine
  Loop
  inCsv.Close
  readCSVIntoArray = rows
End Function