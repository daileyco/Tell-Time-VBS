Dim hour_now
Dim minute_now 
Dim speaks 
Dim speech


hour_now = hour(time)
minute_now = minute(time)

If minute_now = 0 Then
	speaks = "Il est " & hour_now & " heures"
Else
	speaks = "Il est " & hour_now & " heures " & minute_now & " minutes" 
End If

Set speech = CreateObject("sapi.spvoice")




Dim strComputer
Dim TheseProcs
Dim colProcessList
Dim mycount
 

strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")


TheseProcs = Array("zoom", "teams")

mycount = 0

For Each FindProc In TheseProcs
	Set colProcessList = objWMIService.ExecQuery _
		("Select Name from Win32_Process WHERE Name LIKE '" & FindProc & "%'")

	If colProcessList.count>0 Then
		'wscript.echo FindProc & " is running"
		mycount = mycount + 1
	End If
Next



If mycount=0 then
    	speech.Speak speaks
End if


Set objWMIService = Nothing
Set colProcessList = Nothing