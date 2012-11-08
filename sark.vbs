'**************************
'*
'* SARK - Automates a number of routine HOSA tasks
'*
'* author: Adrian Hannah
'* version: 1
'*
'**************************

Dim serv, drac, doors
Dim today, yest, firstDayM, firstDayQ
Dim workstations
Dim fso, folder, files, readFile, readStr, methodStr

serv = "sa96ab1"
drac = "sa96a7e"
doors = "T1W"
workstations = "\\" & serv & "\HOSA$\workstations.txt"

today = Date()
yest = DateAdd("d", -1, today)
firstDayM = DateSerial(Year(Now()),Month(Now()),1)
firstDayQ = DateSerial(Year(Now()),Int((Month(Now())-1)/3)*3+1,1)

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Set yest to the date of the last weekday
Do While Weekday(yest) < 2 Or Weekday(yest) > 6
	yest = DateAdd("d", -1, yest)
Loop

' Set firstDayM to the date of the first weekday of this month
Do While Weekday(firstDayM) < 2 Or Weekday(firstDayM) > 6
	firstDayM = DateAdd("d", 1, firstDayM)
Loop

' Set firstDayQ to the date of the first weekday of this quarter
Do While Weekday(firstDayQ) < 2 Or Weekday(firstDayQ) > 6
	firstDayQ = DateAdd("d", 1, firstDayQ)
Loop

'Daily Scripts
Set folder = fso.GetFolder("daily")
Set files = folder.Files
For Each file In files
	If UCase(fso.GetExtensionName(file.name)) = "VBS" Then
		Set readFile = fso.OpenTextFile("daily/" & file.name, 1, False)
		readStr = readFile.ReadAll
		readFile.close
		ExecuteGlobal readStr
		Wscript.Echo "=========="
	End If
Next

' Weekly Scripts
If Weekday(today) = 2 Then
	Set folder = fso.GetFolder("weekly")
	Set files = folder.Files
	For Each file In files
		If UCase(fso.GetExtensionName(file.name)) = "VBS" Then
			Set readFile = fso.OpenTextFile("weekly/" & file.name, 1, False)
			readStr = readFile.ReadAll
			readFile.close
			ExecuteGlobal readStr
			Wscript.Echo "=========="
		End If
	Next
End If

' Monthly Scripts
If today = firstDayM Then
	Set folder = fso.GetFolder("monthly")
	Set files = folder.Files
	For Each file In files
		If UCase(fso.GetExtensionName(file.name)) = "VBS" Then
			Set readFile = fso.OpenTextFile("monthly/" & file.name, 1, False)
			readStr = readFile.ReadAll
			readFile.close
			ExecuteGlobal readStr
			Wscript.Echo "=========="
		End If
	Next
End If

' Quarterly Scripts
If today = firstDayQ Then
	Set folder = fso.GetFolder("quarterly")
	Set files = folder.Files
	For Each file In files
		If UCase(fso.GetExtensionName(file.name)) = "VBS" Then
			Set readFile = fso.OpenTextFile("quarterly/" & file.name, 1, False)
			readStr = readFile.ReadAll
			readFile.close
			ExecuteGlobal readStr
			Wscript.Echo "=========="
		End If
	Next
End If