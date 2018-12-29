
' #===========================================================#
' #  WINVAP.VBS                                               #
' #===========================================================#
' #  Windows Virtual Access Point (Wi-Fi HotSpot)             #
' #                                                           #
' #       Copyright(C) ZulNs, Gorontalo, October 21'st, 2015  #
' #===========================================================#

Option Explicit

Const TITLE = "Windows Virtual Access Point (Wi-Fi HotSpot) by ZulNs"
Const ADMIN = "~adm"
Const CMD_SETUP = "setup"

Dim param, shell, wsh, ssid, key, msg

If WScript.Arguments.Count > 0 Then
	param = WScript.Arguments.Item(0)
Else
	param = ""
End If

If param <> ADMIN Then
	Set shell = CreateObject("Shell.Application")
	shell.ShellExecute "wscript.exe", Chr(34) + WScript.ScriptFullName + Chr(34) + " " + ADMIN + " " + param, "", "runas", 1
	Set shell = Nothing
	Wscript.Quit
End If

If WScript.Arguments.Count > 1 Then
	param = LCase(WScript.Arguments.Item(1))
Else
	param = ""
End If

Set wsh = CreateObject("WScript.Shell")

If param = "" Then
	If IsStarted() Then
		StopHosting "Do you want to stop hosting now?"
	Else
		StartHosting "Do you want to start hosting now?"
	End If
	ShowHosting
ElseIf param = CMD_SETUP Then
	If IsStarted() Then
		If Not StopHosting("Do you want to stop hosting before setup the hosted network?") Then
			ShowMsg "Setup the hosted network canceled..."
			Set wsh = Nothing
			WScript.Quit
		End If
	End If
	ssid = wsh.ExpandEnvironmentStrings("%COMPUTERNAME%")
	ssid = InputBox("Enter SSID:", TITLE, ssid)
	
	Do
		key = InputBox("Enter network key:", TITLE)
		If key = "" Then
			ShowMsg "Setup the hosted network canceled..."
			Set wsh = Nothing
			WScript.Quit
		End If
		If Len(key) >= 8 Then
			Exit Do
		End If
	Loop
	
	ShowMsg ExecCommand(SetCommand("set") + " mode=allow ssid=" + Chr(34) + ssid + Chr(34) + " key=" + Chr(34) + key + Chr(34))
	StartHosting "Do you want to start hosting now?"
	ShowHosting
Else
	msg = "Unsupported command!!!" + vbCrLf
	msg = msg + vbCrLf
	msg = msg + "Use one of both below commands instead:" + vbCrLf
	msg = msg + vbCrLf
	msg = msg + WScript.ScriptName + vbCrLf
	msg = msg + vbCrLf
	msg = msg + "     to start or stop hosting, or" + vbCrLf
	msg = msg + vbCrLf
	msg = msg + WScript.ScriptName + " setup" + vbCrLf
	msg = msg + vbCrLf
	msg = msg + "     to setup hosted network."
	ShowCriticalMsg msg, True
End If

Set wsh = Nothing
WScript.Quit

'=====================================================
' Subroutines
'=====================================================

Sub ShowMsg(msg)
	MsgBox msg, vbOkOnly, TITLE
End Sub

Sub ShowStyledMsg(msg, style)
	MsgBox msg, vbOkOnly + style, TITLE
End Sub

Sub ShowCriticalMsg(msg, critical)
	Dim style
	If critical Then
		style = vbCritical
	Else
		style = 0
	End If
	ShowStyledMsg msg, style
End Sub

Sub ShowHosting()
	ShowMsg ExecCommand(SetCommand("show"))
End Sub

Function GetResponse(msg)
	GetResponse = MsgBox(msg, vbYesNo + vbQuestion, TITLE) = vbYes
End Function

Function SetCommand(cmd)
	SetCommand = "netsh wlan " + cmd + " hostednetwork"
End Function

Function StartHosting(msg)
	StartHosting = False
	If GetResponse(msg) Then
		Dim temp
		temp = ExecCommand(SetCommand("start"))
		StartHosting = IsStarted()
		ShowCriticalMsg temp, Not StartHosting
	End If
End Function

Function StopHosting(msg)
	StopHosting = False
	If GetResponse(msg) Then
		Dim temp
		temp = ExecCommand(SetCommand("stop"))
		StopHosting = Not IsStarted()
		ShowCriticalMsg temp, Not StopHosting
	End If
End Function

Function IsStarted()
	Dim temp
	temp = LCase(ExecCommand(SetCommand("show")))
	If InStr(temp, "not started") > 0 Then
		IsStarted = False
	Else
		IsStarted = True
	End If
End Function

Function ExecCommand(commandLine)
	Dim exec, stdOut
	Set exec = wsh.Exec(commandLine)
	Set stdOut = exec.StdOut
	While Not stdOut.AtEndOfStream
		ExecCommand = ExecCommand + stdOut.ReadLine + vbCrLf
	Wend
	Set stdOut = Nothing
	Set exec = Nothing
End Function
