' Getter for a specific sd SMART status
' Uses Smartmontools
' Smartmontools folder suggested to be added to %PATH% before script starts
' Outputs JSON with SMART data or health test result
' V 1.2

' CONSTANTS
Const retSuccess   = 0  ' Successful exit
Const retNoParam   = -1 ' Script called without parameters
Const retBadParam  = -2 ' Script called with incorrect parameters

Const LogMaxSize   = 16777216 ' bytes

Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8

Const LogPath      = "C:\Program Files\Zabbix Agent\\Scripts\ScriptData\Logs\smsmartgetter.log"
Const LogPrevPath  = "C:\Program Files\Zabbix Agent\Scripts\ScriptData\Logs\smsmartgetter_prev.log"

Const OutSMART     = "smart"
				   
' VARIABLES        
Set objFSO         = CreateObject("Scripting.FileSystemObject")
OutPath            = "C:\Program Files\Zabbix Agent\Scripts\ScriptData\"
				   
jsonout            = "[{"

' FUNCTIONS
Function FormatNow
	dnow = Now()
	logday = Day(dnow)
	If logday < 10 Then logday = "0" & logday
	logmonth = Month(dnow)
	If logmonth < 10 Then logmonth = "0" & logmonth
	loghour = Hour(dnow)
	If loghour < 10 Then loghour = "0" & loghour
	logminute = Minute(dnow)
	If logminute < 10 Then logminute = "0" & logminute
	logsec = Second(dnow)
	If logsec < 10 Then logsec = "0" & logsec
	FormatNow = logday & "/" & logmonth & "/" & Year(dnow) & " " & _
				loghour & ":" &logminute & ":" & logsec
End Function

Sub LogAddLine(line)
	If objFSO.FileExists(LogPath) Then
		Set objFile = objFSO.GetFile(LogPath)
		If ObjFile.Size < LogMaxSize Then
			Set objFile = Nothing
			Set outputFile = objFSO.OpenTextFile(LogPath, ForAppending, True, -1)
			outputFile.WriteLine(FormatNow & " - " & line)
			outputFile.Close
			Set outputFile = Nothing
		Else
			Set objFile = Nothing
			objFSO.CopyFile LogPath, LogPrevPath, True
			Set outputFile = objFSO.CreateTextFile(LogPath, ForWriting, True)
			outputFile.WriteLine(FormatNow & " - " & line)
			outputFile.Close
			Set outputFile = Nothing
		End If
	Else
		Set outputFile = objFSO.CreateTextFile(LogPath, True, -1)
		outputFile.WriteLine(FormatNow & " - " & line)
		outputFile.Close
		Set outputFile = Nothing
	End If
End Sub

' SCRIPT
LogAddLine "Script started"
If WScript.Arguments.Count > 1 Then
	For I = 0 To WScript.Arguments.Count - 1
		fullArg = fullArg & WScript.Arguments(I) & " "
	Next
	fullArg = Mid(fullArg, 1, Len(fullArg) - 1)
	LogAddLine "Arguments: " & fullArg
	strSMParam = "/dev/" & WScript.Arguments(0)
	argStr = Replace(WScript.Arguments(0), Chr(34), "")
	If Len(argStr) >= 3 Then
		Set objShell = WScript.CreateObject("WScript.Shell")
			Set objExecObject = objShell.Exec("smartctl -a """ & strSMParam & """ -j")
			strOutput = objExecObject.StdOut.ReadAll
			If Instr(strOutput, """severity"": ""error""") = 0 Then
				C = 0
				infoEnd = 0
				infoStart = Instr(strOutput, """device""")
				dataLen = Len(strOutput) - infoStart
				strOutput = Mid(strOutput, infoStart, dataLen)
				outSpl = Split(strOutput, vbCrLf)
				For I = 0 To UBound(outSpl)
					jsonout = jsonout & Trim(Replace(outSpl(I), vbTab, ""))
				Next
				jsonout = Mid(jsonout, 1, Len(jsonout) - 1)
				jsonout = jsonout & "]"
				OutPath = OutPath & WScript.Arguments(0) & " - " & OutSMART & "_" & WScript.Arguments(1) & ".txt"
				Set outFile = objFSO.CreateTextFile(OutPath, True, False)
				outFile.Write jsonout
				outFile.Close
				Set outFile = Nothing
				LogAddLine "Data requested successfully"
				LogAddLine "Script finished"
				Set objFSO = Nothing
				WScript.Echo retSuccess
			Else
				LogAddLine "Incorrect script parameters"
				LogAddLine "Script finished"
				WScript.Echo retBadParam
			End If
			Set objExecObject = Nothing
		Set objShell = Nothing
	Else
		LogAddLine "Incorrect script parameters"
		LogAddLine "Script finished"
		WScript.Echo retBadParam
	End If
Else
	LogAddLine "No script parameters"
	LogAddLine "Script finished"
	WScript.Echo retNoParam
End If