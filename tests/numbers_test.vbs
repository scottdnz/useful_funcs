'Test functions for the "numbers.vbs" library
'Author: Scott Davies (scottdnz)

Sub includeFile(fSpec)
  executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

includeFile "..\lib\common_testing.vbs"
includeFile "..\lib\numbers.vbs"


'Test functions ###############################################################
Sub testisNumberGood()
  goodVals = Array(123, _
    123.5, _
    "123", _
    "123.50", _
    "$123", _
    "$123.50" _
    )
  for each val in goodVals
    AssertEqual True, IsNumeric(val), "testisNumberGood"
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
    AssertNotEqual True, IsNumeric(val), "testisNumberBad"
  next
  AssertNotEqual True, isNumber(emptyVal), "testIsNumberBad"
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

Sub testGetFilteredDblStrg()
  Dim numVals(8,1)
  '0 index is the value, 1 is the expected conversion
  numVals(0,0) = "123.50"
  numVals(0,1) = "123.50"
  
  numVals(1,0) = "$123.50"
  numVals(1,1) = "123.50"
  
  numVals(2,0) = "1,123.50"
  numVals(2,1) = "1123.50"
  
  numVals(3,0) = "$1,123.50"
  numVals(3,1) = "1123.50"
  
  numVals(4,0) = "123"
  numVals(4,1) = "123"
  
  numVals(5,0) = "123a"
  numVals(5,1) = "123"
  
  numVals(6,0) = "abc"
  numVals(6,1) = ""
  
  numVals(7,0) = "-123"
  numVals(7,1) = "-123"
  
  numVals(8,0) = "12#%^3@#^45.00"
  numVals(8,1) = "12345.00"
  
  for i = 0 to Ubound(numVals)
    dblVal = getFilteredDblStrg(numVals(i, 0))
    AssertEqual dblVal, numVals(i, 1), "testGetFilteredDblStrg"
  next
End Sub


'Execute tests ################################################################

testProcs = Array("testIsNumberBad()", _
  "testIsNumberGood()", _
  "testGetFilteredDblStrg()" _
  )
    
runReportTests(testProcs)