'Functions for dealing with numbers


Sub includeFile(fSpec)
  executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

'includeFile "common_testing.vbs"

'Performs a check on a variable/object to determine whether it can be converted
'to a number. Returns a bool.
Function isNumber(val)
  valid = True
  If IsNull(val) or IsEmpty(val) or Not IsNumeric(val) then
    'or (val Is Nothing)
    valid = False
  end if
  isNumber = valid
End Function

'Filters a string so that it only contains permitted characters that represent
'a valid number. Returns a string.
Function getFilteredDblStrg(strgVal)
  allowedChars = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-")
  cleanStrg = ""
  for i = 1 to len(strgVal)
    if inArray(Mid(strgVal, i, 1), allowedChars) Then
      cleanStrg = cleanStrg & Mid(strgVal, i, 1)
    end if
  next
  getFilteredDblStrg = cleanStrg
End Function

Function convMoneyDisplay(strgVal)
  dblVal = FormatNumber(CDbl(strgVal), 2)
  convMoneyDisplay = dblVal
End Function
