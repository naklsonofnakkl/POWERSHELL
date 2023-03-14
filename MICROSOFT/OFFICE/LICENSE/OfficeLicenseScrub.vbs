'*******************************************************************************
' Name: OLicenseCleanup_O16_All.vbs - v 1.09
' Author: Microsoft Customer Support Services
' Copyright (c) Microsoft Corporation
' 
' Removes all licenses for Office 2016 
' from the (Office) Software Protection Platform
'*******************************************************************************

CONST OfficeAppId = "0ff1ce15-a989-479d-af46-f275c6370663"  'Office 2013/2016
Const SKUFILTER = "" 
Const HKLM  = &H80000002

Dim oProductInstances, oWmiLocal, oReg, oWShell, oFso
Dim sQuery
Dim f64, fO64, fQuiet

fQuiet = True

' MAIN
On Error Resume Next
Set oWmiLocal 	= GetObject("winmgmts:\\.\root\cimv2")
Set oReg 		= GetObject("winmgmts:\\.\root\default:StdRegProv")
Set oWShell 	= CreateObject("WScript.Shell")
Set oFso		= CreateObject("Scripting.FileSystemObject")

'Check if we're running as 32 bit process on a 64 bit OS
If InStr(LCase(wscript.path), "syswow64") > 0 Then RelaunchAs64Host

CleanOSPP SKUFILTER
DetectOSBitness
DetectOfficeBitness
ResetOfficeIdentityKey
ResetOfficeUserRegistrationKey
ClearCredmanCache
ClearSCALicCache
' END

'-------------------------------------------------------------------------------
'   ClearSCALicCache
'
'   clear local license cache for SharedComputerActivation 
'-------------------------------------------------------------------------------
Sub ClearSCALicCache
	Dim attr, fld
	Dim sLocalAppData, sCmd, sDelFld
	
	sLocalAppData = oWShell.ExpandEnvironmentStrings("%localappdata%")
	sDelFld = sLocalAppData & "\Microsoft\Office\16.0\Licensing"
	If oFso.FolderExists(sDelFld) Then
		Set fld = oFso.GetFolder(sDelFld)
		'ensure to remove read only flag
		attr = fld.Attributes
		If CBool(attr And 1) Then fld.Attributes = attr And (attr - 1)
		'delete folder
		fld.Delete True
		Set fld = Nothing
		
		'check if removal succeeded. If not try to RD
		If oFso.FolderExists(sDelFld) Then
	        sCmd = "cmd.exe /c rd /s " & chr(34) & sDelFld & chr(34) & " /q"
	        oWShell.Run sCmd, 0, True
        End If
	End If

End Sub 'ClearSCALicCache

'-------------------------------------------------------------------------------
'   ClearCredmanCache
'
'   clear Office credentials from Windows Credentials Manager Cache
'-------------------------------------------------------------------------------
Sub ClearCredmanCache
	Dim oExec, line
	Dim sCmd, sRetVal, sCmdOut, sLine
	Dim arrLines
	
	sCmd = "cmdkey.exe /list:MicrosoftOffice1*"
    Set oExec = oWShell.Exec(sCmd)
    sCmdOut = oExec.StdOut.ReadAll()
    Do While oExec.Status = 0
         WScript.Sleep 100
    Loop
    arrLines = Split(sCmdOut)
	For Each line In arrLines
		If InStr(line, "MicrosoftOffice1") > 0 And Not InStr(line, "MicrosoftOffice1*") > 0 Then 
			sLine = Replace(line, vbCrLf, "")
			sCmd = "cmdkey.exe /delete:" & Trim(sLine)
			sRetVal = oWShell.Run(sCmd, 0, True)
		End If
	Next
End Sub 'ClearCredmanCache

'-------------------------------------------------------------------------------
'   ResetOfficeIdentityKey
'
'   configures the Office Identity key to be reset on next application launch
'-------------------------------------------------------------------------------
Sub ResetOfficeIdentityKey ()
	Dim sSettingsKey, sCount, sRetVal, sCmd
	Dim iCount
	Dim oExec
	
	'remove current user key
	sRetVal = oWShell.Run("REG DELETE HKCU\Software\Microsoft\Office\16.0\Common\Identity /f", 0, True)
	
	'create user settings key to cover other profiles
	sSettingsKey = "SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings"
	If (f64 And fO64) Or (Not f64) Then sSettingsKey = "SOFTWARE\Microsoft\Office\16.0\User Settings"
	
	oReg.CreateKey HKLM, sSettingsKey & "\ResetIdentity"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetIdentity\Delete"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetIdentity\Delete\Software"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetIdentity\Delete\Software\Microsoft"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetIdentity\Delete\Software\Microsoft\Office"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetIdentity\Delete\Software\Microsoft\Office\16.0"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetIdentity\Delete\Software\Microsoft\Office\16.0\Common"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetIdentity\Delete\Software\Microsoft\Office\16.0\Common\Identity"
	
	iCount = 1
	If RegReadDWordValue(HKLM, sSettingsKey & "\ResetIdentity", "Count", sCount) Then iCount = CInt(sCount) + 1
	oReg.SetDWordValue HKLM, sSettingsKey & "\ResetIdentity", "Count", iCount
	oReg.SetDWordValue HKLM, sSettingsKey & "\ResetIdentity", "Order", 1

End Sub 'ResetOfficeIdentityKey

'-------------------------------------------------------------------------------
'   ResetOfficeUserRegistrationKey
'
'   configures the Office Identity key to be reset on next application launch
'-------------------------------------------------------------------------------
Sub ResetOfficeUserRegistrationKey ()
	Dim sSettingsKey, sCount, sRetVal, sCmd
	Dim iCount
	Dim oExec
	
	'remove current user key
	sRetVal = oWShell.Run("REG DELETE HKCU\Software\Microsoft\Office\16.0\Registration /f", 0, True)
	
	'create user settings key to cover other profiles
	sSettingsKey = "SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings"
	If (f64 And fO64) Or (Not f64) Then sSettingsKey = "SOFTWARE\Microsoft\Office\16.0\User Settings"
	
	oReg.CreateKey HKLM, sSettingsKey & "\ResetUserRegistration"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetUserRegistration\Delete"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetUserRegistration\Delete\Software"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetUserRegistration\Delete\Software\Microsoft"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetUserRegistration\Delete\Software\Microsoft\Office"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetUserRegistration\Delete\Software\Microsoft\Office\16.0"
	oReg.CreateKey HKLM, sSettingsKey & "\ResetUserRegistration\Delete\Software\Microsoft\Office\16.0\Registration"
	
	iCount = 1
	If RegReadDWordValue(HKLM, sSettingsKey & "\ResetUserRegistration", "Count", sCount) Then iCount = CInt(sCount) + 1
	oReg.SetDWordValue HKLM, sSettingsKey & "\ResetUserRegistration", "Count", iCount
	oReg.SetDWordValue HKLM, sSettingsKey & "\ResetUserRegistration", "Order", 1

End Sub 'ResetOfficeUserRegistrationKey

'-------------------------------------------------------------------------------
'   DetectOfficeBitness
'
'   detect bitness of Office
'-------------------------------------------------------------------------------
Sub DetectOfficeBitness ()
	Dim sOPlatform, sInstallRootPath

	fO64 = False
	If Not f64 Then Exit Sub
	If RegReadStringValue(HKLM, "SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "platform", sOPlatform) Then
		fO64 = (sOPlatform = "x64")
		Exit Sub
	End If
	If RegReadStringValue(HKLM, "SOFTWARE\Microsoft\Office\ClickToRun\propertyBag", "Platform", sOPlatform) Then
		fO64 = (sOPlatform = "x64")
		Exit Sub
	End If
	If RegReadStringValue(HKLM, "SOFTWARE\Wow6432Node\Microsoft\Office\Common\InstallRoot", "Path", sInstallRootPath) Then
		'fO64 = Not (InStr(sInstallRootPath,"(x86)") > 0)
		fO64 = False
		Exit Sub
	End If
	If RegReadStringValue(HKLM, "SOFTWARE\Microsoft\Office\Common\InstallRoot", "Path", sInstallRootPath) Then
		'fO64 = Not (InStr(sInstallRootPath,"(x86)") > 0)
		fO64 = True
		Exit Sub
	End If

End Sub 'DetectOfficeBitness

'-------------------------------------------------------------------------------
'   DetectOSBitness
'
'   detect bitness of the operating system
'-------------------------------------------------------------------------------
Sub DetectOSBitness ()
	Dim ComputerItem, item
	
    Set ComputerItem = oWmiLocal.ExecQuery("Select * from Win32_ComputerSystem")
    For Each item In ComputerItem
        f64 = Instr(Left(item.SystemType, 3), "64") > 0
    Next
End Sub 'DetectOSBitness

'-------------------------------------------------------------------------------
'   CleanOSPP
'
'   unpkeys the licenses from OSPP
'-------------------------------------------------------------------------------
Sub CleanOSPP (sFilter)

	' Initialize the software protection platform object with a filter on Office 2013/2016 products
	If GetVersionNT > 601 Then
	    Set oProductInstances = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Description, Name, ProductKeyID FROM SoftwareLicensingProduct WHERE ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL")
	Else
	    Set oProductInstances = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Description, Name, ProductKeyID FROM OfficeSoftwareProtectionProduct WHERE ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL")
	End If
	
	' Remove all licenses
	For Each pi in oProductInstances
	    If Not fQuiet Then WScript.Echo pi.Name
	    If NOT IsNull(pi) Then
	        If InStr(pi.Name, sFilter) > 0 Or sFilter = "" Then
		        If Not fQuiet Then WScript.Echo " -> Uninstall ProductKey: " & pi.ProductKeyID
		        pi.UninstallProductKey(pi.ProductKeyID)
	        End If
	    End If
	Next 'pi
End Sub 'CleanOSPP

'-------------------------------------------------------------------------------
'   GetVersionNT
'
'   Calculate the VerionNT number as integer 
'-------------------------------------------------------------------------------
Function GetVersionNT ()

    Dim sOsVersion, sOsItem
    Dim arrVersion
    Dim qOS

    Set qOS = oWmiLocal.ExecQuery( "Select * from Win32_OperatingSystem")
    For Each oOsItem in qOS 
        sOsVersion = oOsItem.Version
    Next
    arrVersion = Split( sOsVersion, GetDelimiter( sOsVersion))
    GetVersionNT = CInt( arrVersion( 0)) * 100 + CInt( arrVersion( 1))
End Function

'-------------------------------------------------------------------------------
'   GetDelimiter
'
'   Returns the delimiter in a version string 
'-------------------------------------------------------------------------------
Function GetDelimiter (sVersion)
    Dim iCnt, iAsc

    GetDelimiter = " "
    For iCnt = 1 To Len(sVersion)
        iAsc = Asc(Mid(sVersion, iCnt, 1))
        If Not (iASC >= 48 And iASC <= 57) Then 
            GetDelimiter = Mid(sVersion, iCnt, 1)
            Exit Function
        End If
    Next 'iCnt
End Function

'-------------------------------------------------------------------------------
'   RegReadDWordValue
'
'   Check if a string value exists and return on zero if not
'-------------------------------------------------------------------------------
Function RegReadDWordValue(hDefKey, sSubKeyName, sName, sValue)
    Dim RetVal
    
    RetVal = oReg.GetDWORDValue(hDefKey, sSubKeyName, sName, sValue)
    RegReadDWordValue = (RetVal = 0)
    
End Function 'RegReadDWordValue

'-------------------------------------------------------------------------------
'   RegReadStringValue
'
'   Check if a string value exists and return on zero if not
'-------------------------------------------------------------------------------
Function RegReadStringValue(hDefKey, sSubKeyName, sName, sValue)
    Dim RetVal
    
    RetVal = oReg.GetStringValue(hDefKey, sSubKeyName, sName, sValue)
    RegReadStringValue = (RetVal = 0)
    
End Function 'RegReadSringValue

'-------------------------------------------------------------------------------
'   RelaunchAs64Host
'
'   Relaunch self with 64 bit CScript host
'-------------------------------------------------------------------------------
Sub RelaunchAs64Host
    Dim Argument, sCmd
    Dim fQuietRelaunch

    fQuietRelaunch = False
    sCmd = Replace(LCase(wscript.Path), "syswow64", "sysnative") & "\cscript.exe " & Chr(34) & WScript.scriptFullName & Chr(34)
    If fQuiet Then fQuietRelaunch = True
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            sCmd = sCmd  &  " " & chr(34) & Argument & chr(34)
            Select Case UCase(Argument)
            Case "/Q", "/QUIET"
                fQuietRelaunch = True
            End Select
        Next 'Argument
    End If
    sCmd = sCmd & " /ChangedHostBitness"
    If fQuietRelaunch Then
        sCmd = Replace (sCmd, "\cscript.exe", "\wscript.exe")
        Wscript.Quit CLng(oWShell.Run (sCmd, 0, True))
    Else
        Wscript.Quit CLng(oWShell.Run (sCmd, 1, True))
    End If

End Sub 'RelaunchAs64Host
