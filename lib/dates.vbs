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

Function removeDaySuffix(dayWSuffix)
  Dim suffixes 'Array
  Dim sfx  'Strings
  'Strip the suffix from dayWSuffix
  suffixes = Array("st", "th", "nd", "rd")
  for each sfx in suffixes
    dayWSuffix = Replace(dayWSuffix, sfx, "")
  next
  removeDaySuffix = dayWSuffix
End Function

'Converts two string parameters, e.g. "15th" & "February 2014" to a date string
'in the format dd/mm/yyyy
Function convMonthNameDateToFmtStrg(dayWSuffix, monthNameYear)
  dayWithoutSfx = removeDaySuffix(dayWSuffix)
  dayWithoutSfx = addLeadingZero(dayWithoutSfx)
  monthNum = Month(monthNameYear)
  monthNum = addLeadingZero(monthNum)
  yearNum = Year(monthNameYear)
  convMonthNameDateToFmtStrg = dayWithoutSfx & "/" & monthNum & "/" & yearNum
End Function

'Takes a 2-d array of dates information and writes it to a CSV file.
Function writeCSV(dtsArr)
  Dim objFSO 'Object
  Dim objFile 'Object
  Dim outFile 'String
  Dim csvArr 'Array
  Dim i 'Int
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  outFile = ".\output_" & getDtForFName(Now) & ".csv"
  Set objFile = objFSO.CreateTextFile(outFile, True)
  'Write the Header row
  objFile.WriteLine("Current date,7 WD Future") 
  for i = 0 to UBound(dtsArr)
    csvArr = Array(dtsArr(i, 0), _
                   dtsArr(i, 1))
    objFile.WriteLine(Join(csvArr, ","))
  next
  objFile.Close
  'wscript.echo "File " & outFile & " written."
  writeCSV = outFile
End Function

Function isValidDate(dt)
  Dim re 'Object
  Dim match 'Object
  Dim numMatches 'Array-like object
  Dim cntr, isValid 'Int
  Set re = New RegExp
  re.Global = True
  'Must match the date pattern d/m/yyyy or dd/mm/yyyy
  re.Pattern = "^\d{1,2}/\d{1,2}/\d{4}$"
  set numMatches = re.Execute(dt)  
  cntr = 0
  For Each match In numMatches
    cntr = cntr + 1
  Next
  If cntr = 1 Then
    isValid = True
  Else
    isValid = False
  End If
  isValidDate = isValid 
End Function

Function isValidTime(tm, hoursFmt)
  Dim re 'Object
  Dim match 'Object
  Dim numMatches 'Array-like object
  Dim cntr, isValid, hr, mn 'Int
  if hoursFmt = "12"
    maxHours = 12
  else
    maxHours = 24
  end if
  isValid = True
  re.Pattern = "^\d{1,2}\:\d{1,2}$"
  set numMatches = re.Execute(tmStrg)  
  cntr = 0
  For Each match In numMatches
    cntr = cntr + 1
  Next
  hr = CInt(splitted(0))
  mn = CInt(splitted(1))
  if hr < 0 or hr > maxHours Then
    'Hour is too low or too high
    isValid = False
  end if 
  if mn < 0 or mn > 59 Then
    'Minutes is too low or too high
    isValid = False
  End If
End Function


'monthsArr = Array("January", "February", "March", "April", "May", "June", _
'"July", "August", "September", "October", "November", "December")

