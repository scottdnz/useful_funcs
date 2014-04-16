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
  dtObj = CDate("2014-11-27 16:40:32") 'gives a datetime
  dtForName = getDtForFName(dtObj)
  AssertEqual dtForName, "20141127_164032", "testGetDtForFName"
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

Sub testRemoveDaySuffix()
  Dim dtsArray(3,1)
  dtsArray(0,0) = "1st"
  dtsArray(0,1) = "1"
  dtsArray(1,0) = "2nd"
  dtsArray(1,1) = "2"
  dtsArray(2,0) = "3rd"
  dtsArray(2,1) = "3"
  dtsArray(3,0) = "4th"
  dtsArray(3,1) = "4"
  for i = 0 to Ubound(dtsArray)
    dayWSuffix = dtsArray(i, 0)
    expectedDayWithoutSfx = dtsArray(i, 1)
    dayWithoutSfx = removeDaySuffix(dayWSuffix)
    'Wscript.echo dayWithoutSfx & ", " & expectedDayWithoutSfx
    AssertEqual dayWithoutSfx, expectedDayWithoutSfx, "testRemoveDaySuffix"
  next
End Sub

Sub testConvMonthNameDateToFmtStrg()
  Dim dtsArray(3,2)
  dtsArray(0,0) = "1st"
  dtsArray(0,1) = "February 2014"
  dtsArray(0,2) = "01/02/2014"
  dtsArray(1,0) = "22nd"
  dtsArray(1,1) = "March 2013"
  dtsArray(1,2) = "22/03/2013"
  dtsArray(2,0) = "13th"
  dtsArray(2,1) = "December 2013"
  dtsArray(2,2) = "13/12/2013"
  dtsArray(3,0) = "23rd"
  dtsArray(3,1) = "July 2014"
  dtsArray(3,2) = "23/07/2014"
  for i = 0 to Ubound(dtsArray)
    dayWSuffix = dtsArray(i, 0)
    monthNameYear = dtsArray(i, 1)
    expectedDtFmtStrg = dtsArray(i, 2)
    dtFmtStrg = convMonthNameDateToFmtStrg(dayWSuffix, monthNameYear)
    'Wscript.echo dayWSuffix & ", " & monthNameYear & ", " & dtFmtStrg
    AssertEqual dtFmtStrg, expectedDtFmtStrg, "testConvMonthNameDateToFmtStrg"
  next
End Sub

Sub testIsValidDatePositives()
  Dim dtStrgArr
  Dim dt
  dtStrgArr = Array("1/1/2014", "01/1/2014", "1/01/2014", "01/01/2014")
  for each dt in dtStrgArr
    AssertEqual isValidDate(dt), 1, "testIsValidDatePositives"
    'wscript.echo dt & ": " & isValidDate(dt)
  next
End Sub

Sub testIsValidDatePositives()
  Dim dtStrgArr
  Dim dt
  dtStrgArr = Array("1/1/2014", "01/1/2014", "1/01/2014", "01/01/2014")
  for each dt in dtStrgArr
    AssertEqual isValidDate(dt), True, "testIsValidDatePositives"
    'wscript.echo dt & ": " & isValidDate(dt)
  next
End Sub

Sub testIsValidDateNegatives()
  Dim dtStrgArr
  Dim dt
  dtStrgArr = Array("111/1/2014", "01/111/2014", "1/01/201", "/01/2014", _
  "1/1/2014 05:14:33pm")
  for each dt in dtStrgArr
    AssertEqual isValidDate(dt), False, "testIsValidDatePositives"
  next
End Sub

'Sub testIsValidTime()
'  'IsValidTime(tm, hoursFmt)
'End Sub

Sub testGetDiffDates()
  Dim datesArr(7,2)
  Dim diffMsg
  Dim d1, d2
  datesArr(0,0) = "17/03/2014"
  datesArr(0,1) = "18/08/2008"
  datesArr(0,2) = "5 year(s), 6 month(s) & 14 day(s)"
  
  datesArr(1,0) = "17/03/2014"
  datesArr(1,1) = ""
  datesArr(1,2) = "Could not be calculated"
  
  datesArr(2,0) = ""
  datesArr(2,1) = "18/08/2008"
  datesArr(2,2) = "Could not be calculated"
  
  datesArr(3,0) = ""
  datesArr(3,1) = ""
  datesArr(3,2) = "Could not be calculated"
  
  datesArr(4,0) = "17/03/2014"
  datesArr(4,1) = "11/03/2013"
  datesArr(4,2) = "1 year(s), 0 month(s) & 6 day(s)"
  
  datesArr(5,0) = "17/03/2014"
  datesArr(5,1) = "17/03/2014"
  datesArr(5,2) = "0 year(s), 0 month(s) & 0 day(s)"
  
  datesArr(6,0) = "11/03/2014"
  datesArr(6,1) = "01/12/1976"
  datesArr(6,2) = "37 year(s), 3 month(s) & 9 day(s)"
  
  datesArr(7,0) = "17/03/2008"
  datesArr(7,1) = "11/03/2014"
  datesArr(7,2) = "5 year(s), 11 month(s) & 8 day(s)"
  
  for i = 0 to 7
    d1 = datesArr(i, 0)
    d2 = datesArr(i, 1)
    expectedDiffMsg = datesArr(i, 2)
    diffMsg = getDiffDates(d1, d2)
    wscript.echo "d1: " & datesArr(i, 0) & ", dt2: " & datesArr(i, 1)
    wscript.echo "diffMsg: " & diffMsg
    wscript.echo "***"
    AssertEqual diffMsg, expectedDiffMsg, "testGetDiffDates"
  next
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
                  "testWriteCSV()", _
                  "testRemoveDaySuffix()", _
                  "testConvMonthNameDateToFmtStrg()", _
                  "testIsValidDatePositives()", _
                  "testIsValidDateNegatives()", _
                  "testGetDiffDates()" _
                  )
runReportTests(testProcs)