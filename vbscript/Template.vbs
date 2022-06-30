'----------------------------------------------------------------------
'	<Title of Software> Installation Script
'	Last Revised: 2014-12-16
'	Revised By:  Lee Miller
'----------------------------------------------------------------------

On Error Resume Next

'-[Declare Constants]--------------------------------------------------

Const OverwriteExisting = True
Const WaitOnReturn = True
Const DisplayInteractiveMessage = True
Const MsgSoftwareTitle = "<Title of Software>"

'-[Declare Variables]--------------------------------------------------
Dim installDone, OSType, strAlreadyInstalled, strContinue, strIsInstalled, strIsMetric, strIsRunning, strQuoted, strVersion
Dim objFSO, objShell, objNetwork 

'-[Create System Objects]----------------------------------------------

Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objShell=CreateObject("Wscript.Shell")
Set objNetwork=CreateObject("Wscript.Network")

'-[Display Message]---------------------------------------------------------------

If DisplayInteractiveMessage = True Then
	Msg = "This will install " & MsgSoftwareTitle & "." & vbCrLf & vbCrLf
	Msg = msg & "The installation takes about 5 minutes and is silent with the exception "
	Msg = msg & "of periodic progress bars that will be displayed." & vbCrLf & vbCrLf
	Msg = msg & "You will be notified when it completes." & vbCrLf & vbCrLf
	Msg = msg & "Please close Navisworks if it is open and then click OK to begin."
	strContinue = MsgBox (msg, 65, "New Software Install/Update Notification")
End If
If strContinue = 2 Then
	WScript.Quit (8)
End If

'Check if the application is running
strIsInstalled = IsInstalled (MsgSoftwareTitle)
If strIsInstalled = False Then
	strIsRunning = IsRunning ("Roamer.exe")
	If strIsRunning = True Then
		If DisplayInteractiveMessage = True Then
			Msg = "<Title of Software> is running so the installation cannot proceed." & vbCrLf & vbCrLf
			Msg = msg & "Please close Navisworks and start the installation again."
			MsgBox msg, 64, "New Software Install/Update Notification"
		End If
		WScript.Quit (4)
	End If
End If	
	
'-[Preinstallation Tasks]---------------------------------------------------------------


'-[MAIN]---------------------------------------------------------------

'Install <Title of Software>
strIsInstalled = IsInstalled (MsgSoftwareTitle)
If strIsInstalled = False Then
	strQuoted = Chr(34) & "\\GROUP\HOK\FWR\RESOURCES\HSD\NAVISWORKS2014-SCCM\Addins\Exporters_R1_2014\Img\setup.exe" & Chr(34) & "  /W /QB /I \\GROUP\HOK\FWR\RESOURCES\HSD\NAVISWORKS2014-SCCM\Addins\Exporters_R1_2014\Img\Navisworks2014Exporters.ini /language en-us"
	objShell.Run strQuoted, 0, WaitOnReturn
Else
	If DisplayInteractiveMessage = True Then
		Msg = MsgSoftwareTitle & " is already installed." & vbCrLf & vbCrLf
		MsgBox msg, 64, "New Software Install/Update Notification"
	End If
End If
On Error Goto 0

'Confirm Installation
strIsInstalled = IsInstalled (MsgSoftwareTitle)
If strIsInstalled = False Then
	If DisplayInteractiveMessage = True Then
		Msg = "The " & MsgSoftwareTitle & " installation failed." & vbCrLf & vbCrLf
		MsgBox msg, 64, "New Software Install/Update Notification"
	End If
	WScript.Quit (14)
End If
If strIsInstalled = True Then
	If DisplayInteractiveMessage = True Then
		Msg = "The " & MsgSoftwareTitle & " installation completed successfully." & vbCrLf & vbCrLf
		MsgBox msg, 64, "New Software Install/Update Notification"
	End If
	WScript.Quit (0)
End If

On Error Goto 0

'======================================================================
'This function accepts a software title and returns true if found. 
'This is a placeholder function that calls on IsInstalled32 and 
'IsInstalled64 to perform the check toward the registry.
'======================================================================
Function IsInstalled(strSoftwareTitle)
	On Error Resume Next
	Dim strIsInstalled64
	strIsInstalled64 = IsInstalled64(strSoftwareTitle)	
	If (strIsInstalled64 = True) Then 
		IsInstalled = True
	Else
		IsInstalled = False
	End If 
	On Error Goto 0	
End Function

'======================================================================
'This function accepts a 64-bit software title and returns True if it is
'installed, otherwise it returns False
'======================================================================
Function IsInstalled64 (strSoftwareTitle)
	Const HKLM = &H80000002
	Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
	Dim strInstalled
	strInstalled = false
	
	objCtx.Add "__ProviderArchitecture", 64
	objCtx.Add "__RequiredArchitecture", TRUE
	Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
	Set objServices = objLocator.ConnectServer("","root\default","","",,,,objCtx)
	Set objStdRegProv = objServices.Get("StdRegProv") 
	
	' Use ExecMethod to call the GetStringValue method
	Set Inparams = objStdRegProv.Methods_("EnumKey").Inparameters
	Inparams.Hdefkey = HKLM
	Inparams.Ssubkeyname = "Software\Microsoft\Windows\CurrentVersion\Uninstall\" 
	Set Outparams = objStdRegProv.ExecMethod_("EnumKey", Inparams,,objCtx) 
	For Each strSubKey In Outparams.snames 
		Set Inparams = objStdRegProv.Methods_("GetStringValue").Inparameters
		Inparams.Hdefkey = HKLM
		Inparams.Ssubkeyname = "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & strSubKey
		Inparams.Svaluename = "DisplayName"
		set Outparams = objStdRegProv.ExecMethod_("GetStringValue", Inparams,,objCtx) 
			if ("" & Outparams.sValue) = "" then
			'wscript.echo strSubKey
			Else
				'wscript.echo Outparams.SValue
				If strSoftwareTitle = Outparams.Svalue Then
					strInstalled = True 
				End If 
			End iF 
	Next
	IsInstalled64 = strInstalled
End Function

'======================================================================
'This function accepts a process name and returns True if it is running,
'otherwise it returns False
'======================================================================
Function IsRunning (StrProcessName)
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcesses = objWMIService.ExecQuery _
	    ("Select * from Win32_Process Where Name = '" & strProcessName & "'")
	If colProcesses.Count = 0 Then
	   IsRunning = False
	Else
	   IsRunning = True
	End If
End Function
