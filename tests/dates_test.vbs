'Test functions for the "dates.vbs" library
'Author: Scott Davies (scottdnz)

Sub includeFile(fSpec)
  executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

includeFile "..\lib\common_testing.vbs"
includeFile "..\lib\dates.vbs"


'Test functions ###############################################################

Sub testAddLeadingZeroPositive()
  Dim intVal
  Dim expectedStrg, strVal
  intVal = 9
  strVal = addLeadingZero(intVal)
  expectedStrg = "09"
  AssertEqual strVal, expectedStrg, "testAddLeadingZeroPositive"
End Sub

Sub testAddLeadingZeroNegative()
  Dim intVal
  Dim expectedStrg, strVal
  intVal = 11
  strVal = addLeadingZero(intVal)
  expectedStrg = "11"
  AssertEqual strVal, expectedStrg, "testAddLeadingZeroNegative"
End Sub
  
Sub testGetNiceShortDatePlain()
  Dim dtObj
  Dim expectedStrg, shortDt
  dtObj = DateSerial(2013, 11, 27)
  shortDt = getNiceShortDate(dtObj)
  expectedStrg = "27/11/2013"
  AssertEqual shortDt, expectedStrg, "testGetNiceShortDatePlain"
End Sub

Sub testGetNiceShortDateZeros()
  Dim dtObj
  Dim expectedStrg, shortDt
  dtObj = DateSerial(2013, 1, 1)
  shortDt = getNiceShortDate(dtObj)
  expectedStrg = "01/01/2013"
  AssertEqual shortDt, expectedStrg, "testGetNiceShortDateZeros"
End Sub

Sub testGetDtForFName()
  Dim dtForName
  Dim dtObj
  dtObj = CDate("2013-11-27 16:40:32") 'gives a datetime
  dtForName = getDtForFName(dtObj)
  AssertEqual dtForName, "20131127_164032", "testGetDtForFName"
End Sub

Sub testGetDtObjFromStrg()
  Dim dtStrg
  Dim dtObj
  dtStrg = "04/12/2013"
  dtObj = getDtObjFromStrg(dtStrg)
  AssertEqual TypeName(dtObj), "Date", "testGetDtObjFromStrg"
End Sub

Sub testGetDtObjFromStrgValues()
  Dim dtStrg
  Dim dtObj
  dtStrg = "04/12/2013"
  dtObj = getDtObjFromStrg(dtStrg)
  AssertEqual Day(dtObj), 4, "testGetDtObjFromStrgValues"
  AssertEqual Month(dtObj), 12, "testGetDtObjFromStrgValues"
  AssertEqual Year(dtObj), 2013, "testGetDtObjFromStrgValues"
End Sub

Sub testGet7WDFuture()
  dtObj = DateSerial(2013, 12, 4)
  expectedDt = DateSerial(2013, 12, 13)
  wD7FutureDt = get7WDFuture(dtObj)
  AssertEqual Year(wD7FutureDt), Year(expectedDt), "testGet7WDFuture"
  AssertEqual Month(wD7FutureDt), Month(expectedDt), "testGet7WDFuture"
  AssertEqual Day(wD7FutureDt), Day(expectedDt), "testGet7WDFuture"
End Sub

Sub testWriteCSV()
  Dim dtsArray(1,1)
  Dim cntr
  dtsArray(0,0) = "05/12/2013"
  dtsArray(0,1) = "17/12/2013"
  dtsArray(1,0) = "06/12/2013"
  dtsArray(1,1) = "18/12/2013"
  outFile = writeCSV(dtsArray)
  'Read the CSV file that was written & check it
  Set inCsvObj = CreateObject("Scripting.FileSystemObject") 
  Set inCsv = inCsvObj.OpenTextFile(outFile, "1", True)
  cntr = 0
  inCsv.ReadLine    'Read & ignore the header line
  Do While Not inCsv.AtEndOfStream
    curLine = inCsv.ReadLine
    rowItems = Split(curLine, ",")   
    AssertEqual rowItems(0), dtsArray(cntr, 0), "testWriteCSV"
    AssertEqual rowItems(1), dtsArray(cntr, 1), "testWriteCSV"
    cntr = cntr + 1
  Loop
  inCsv.Close
  'Clean up
  Set obj = CreateObject("Scripting.FileSystemObject") 'Calls the File System Object  
  obj.DeleteFile(outFile) 
End Sub

'Execute tests ################################################################

testProcs = Array("testAddLeadingZeroPositive()", _
                  "testAddLeadingZeroNegative()", _
                  "testGetNiceShortDatePlain()", _
                  "testGetNiceShortDateZeros()", _                  
                  "testGetDtForFName()", _
                  "testGetDtObjFromStrg()", _
                  "testGetDtObjFromStrgValues()", _
                  "testGet7WDFuture()", _
                  "testWriteCSV()" _
                  )
runReportTests(testProcs)