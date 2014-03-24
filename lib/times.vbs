'Test times

tmsArr = Array("1:03", "12:03", "14:21", "00:00", "0:01", "23:59", "01:03", ":05", "1:0", "1a:05", "12:a5", "1:1a", "aa:aa", "aa:1", ":", "0305", "abcd", "24:01", "-3:60", "03:60")

Set re = New RegExp
re.Global = True
re.Pattern = "^\d{1,2}\:\d{2}$"

for each tmStrg in tmsArr
  set numMatches = re.Execute(tmStrg)  
  cntr = 0
  For Each match In numMatches
    cntr = cntr + 1
  Next
  if cntr <> 1 Then
    wscript.echo "'" & tmStrg & "' does not fit the pattern hh:mm"  
  else
    splitted = Split(tmStrg, ":")
    if Ubound(splitted) <> 1 Then
      wscript.echo "Hour or minutes missing from: " & tmStrg
    End If
    
    hr = CInt(splitted(0))
    mn = CInt(splitted(1))
    if hr < 0 or hr > 23 Then
      wscript.echo "Hour is too low or too high in: " & tmStrg
    end if 
    if mn < 0 or mn > 59 Then
      wscript.echo "Minutes is too low or too high in: " & tmStrg
    End If
  end if
  
next

wscript.echo "**************************************************************"

