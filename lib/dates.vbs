'Useful date functions for VB Script.
'Author: Scott Davies (scottdnz)


'Takes an integer and adds a leading zero to its string value if under 10.
'Returns a string. 
Function addLeadingZero(intVal)
  Dim strVal 'String(s)
  strVal = CStr(intVal)
  if intVal < 10 then
    strVal = "0" & intVal
  end if
  addLeadingZero = strVal
End Function

'Takes a Date object and returns a formatted string, i.e. dd/mm/yyyy.
Function getNiceShortDate(dt)
  Dim curDay, curMonth 'String(s)
  curDay = addLeadingZero(Day(dt))
  curMonth = addLeadingZero(Month(dt))
  getNiceShortDate = curDay & "/" & curMonth & "/" & Year(dt)
End Function

'Gets information from the current DateTime and returns a formatted string based
'on it, i.e. yyyymmdd_HHMMSS format.
Function getDtForFName(nowVal)
  Dim curMonth, curDay, curHour, curMinute, curSecond 'String(s)
  curMonth = addLeadingZero(Month(nowVal))
  curDay = addLeadingZero(Day(nowVal))
  curHour = addLeadingZero(Hour(nowVal))
  curMinute = addLeadingZero(Minute(nowVal))
  curSecond = addLeadingZero(Second(nowVal))
  getDtForFName = Year(Now) & curMonth & curDay & "_" & curHour & curMinute & _
  curSecond
End Function

'Takes a formatted date string in the format dd/mm/yyy, and returns a Date
'object based on it.
Function getDtObjFromStrg(dtStrg)
  Dim splitted 'Array(s)
  Dim curDt 'Date(s)
  splitted = Split(dtStrg, "/")  
  curDt = DateSerial(CInt(splitted(2)), CInt(splitted(1)), CInt(splitted(0)))
  getDtObjFromStrg = curDt
End Function

'Takes a specific Date object, and returns a new Date seven working days in the
'future.
Function get7WDFuture(specDate)
  Dim sevenWDFuture, lowest 'Date(s)
  Dim numDays 'Integer(s)
  Dim wday 'String(s)
  sevenWDFuture = specDate
  numDays = 0
  while numDays < 7
    sevenWDFuture = dateadd("d", 1, sevenWDFuture)
    wday = weekdayname(weekday(sevenWDFuture, 1), true, 1)
    if not (wday = "Sat" or wday = "Sun") then 
      numDays = numDays + 1
    end if
  Wend  
  lowest = DateSerial(Year(sevenWDFuture), Month(sevenWDFuture), Day(sevenWDFuture))
  get7WDFuture = getNiceShortDate(lowest)
End Function

'Takes an array of dates information and writes it to a CSV file.
Sub writeCSV(dtsArr)
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  outFile = ".\output_" & getDtForFName & ".csv"
  Set objFile = objFSO.CreateTextFile(outFile, True)
  'Write the Header row
  objFile.WriteLine("Current date,7 WD Future") 
  
  for i = 0 to UBound(dtsArr)
    csvArr = Array(dtsArr(i, 0), _
                   dtsArr(i, 1))
    objFile.WriteLine(Join(csvArr, ","))
  next
  
  objFile.Close
  wscript.echo "File " & outFile & " written."
End Sub

Sub AssertEqual(given, expected, procName)
  if not given = expected then
    Err.Raise vbObjectError + 99999, , procName & " failed"
  end if
End Sub

Sub AssertGreater(given, expected, procName)
  if not given > expected then
    Err.Raise vbObjectError + 99999, , procName & " failed"
  end if
End Sub

  
'Tests ########################################################################

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
  dtObj = CDate("2013-11-27 16:40:32") 'gives a datetime
  dtForName = getDtForFName(dtObj)
  AssertEqual dtForName, "20131127_164032", "testGetDtForFName"
End Sub

Sub testGetDtObjFromStrg()
  'Placeholder
  'getDtObjFromStrg
End Sub

Sub testGet7WDFuture()
  'Placeholder
  'get7WDFuture
End Sub

Sub testWriteCSV()
  'Placeholder
  'writeCSV(dtsArr)
End Sub


'Execute tests ################################################################
cntr = 0
testProcs = Array("testAddLeadingZeroPositive()", _
                  "testAddLeadingZeroNegative()", _
                  "testGetNiceShortDatePlain()", _
                  "testGetNiceShortDateZeros()", _                  
                  "testGetDtForFName()" _
                  )
                  
For each testProc in testProcs
  Execute testProc
  cntr = cntr + 1
Next
wscript.echo cntr & " tests succesfully passed."
