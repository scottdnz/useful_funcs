'Functions for dealing with numbers


Sub includeFile(fSpec)
  executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

includeFile "common_testing.vbs"


Function IsNumber(val)
  valid = True
  If IsNull(val) or IsEmpty(val) then
    'or (val Is Nothing)
    valid = False
  end if
  If valid Then
    If Not IsNumeric(val) Then
      valid = False
    End If
  End If
  IsNumber = valid
End Function


'Test functions ###############################################################
Sub testIsNumberGood()
  goodVals = Array(123, _
    123.5, _
    "123", _
    "123.50", _
    "$123", _
    "$123.50" _
    )
  for each val in goodVals
    AssertEqual True, IsNumeric(val), "testIsNumberGood"
  next
End Sub

Sub testIsNumberBad()
  Dim emptyVal
  badVals = Array("abc", _
    "123a", _
    DateSerial(2013, 12, 5), _
    "", _
    Null _
    )
  for each val in badVals
    AssertNotEqual True, IsNumeric(val), "testIsNumberBad"
  next
  AssertNotEqual True, IsNumber(emptyVal), "testIsNumberBad"
  'wscript.echo IsNumeric(emptyVal)
End Sub

Sub testIsNumberGood()
  goodVals = Array(123, _
    123.5, _
    "123", _
    "123.50", _
    "$123", _
    "$123.50" _
    )
  for each val in goodVals
    AssertEqual True, IsNumeric(val), "testIsNumberGood"
  next
End Sub

'Execute tests ################################################################

testProcs = Array("testIsNumberBad()", _
  "testIsNumberGood()" _
  )
runReportTests(testProcs)

Dim emptyVal
'val = emptyVal

If IsNull(emptyVal) or IsEmpty(emptyVal) then
  wscript.echo "not"
end if

'wscript.echo CDbl("$1.50")