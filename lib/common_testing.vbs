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