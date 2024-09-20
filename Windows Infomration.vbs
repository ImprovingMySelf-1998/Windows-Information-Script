Const HKEY_LOCAL_MACHINE = &H80000002

' Get the Product Key
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
strValueName = "DigitalProductId"
strComputer = "."
Dim iValues()

Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
  strComputer & "\root\default:StdRegProv")
oReg.GetBinaryValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, iValues

Dim arrDPID
arrDPID = Array()
For i = 52 To 66
    ReDim Preserve arrDPID(UBound(arrDPID) + 1)
    arrDPID(UBound(arrDPID)) = iValues(i)
Next

Dim arrChars
arrChars = Array("B", "C", "D", "F", "G", "H", "J", "K", "M", "P", "Q", "R", "T", "V", "W", "X", "Y", "2", "3", "4", "6", "7", "8", "9")

strProductKey = ""
For i = 24 To 0 Step -1
    k = 0
    For j = 14 To 0 Step -1
        k = k * 256 Xor arrDPID(j)
        arrDPID(j) = Int(k / 24)
        k = k Mod 24
    Next
    strProductKey = arrChars(k) & strProductKey
    If i Mod 5 = 0 And i <> 0 Then strProductKey = "-" & strProductKey
Next

' Get OS Information
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOperatingSystem In colOperatingSystems
    strOS = objOperatingSystem.Caption
    strVersion = objOperatingSystem.Version
    strBuild = objOperatingSystem.BuildNumber
    strSerial = objOperatingSystem.SerialNumber
    strRegistered = objOperatingSystem.RegisteredUser
Next

' Get Computer Manufacturer and Model
Set colComputers = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer In colComputers
    strManufacturer = objComputer.Manufacturer
    strModel = objComputer.Model
Next

' Determine License Type
Dim strLicenseType
Dim oemKeyPath
oemKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\OEMInformation"

On Error Resume Next
Dim oemValue
oReg.GetStringValue HKEY_LOCAL_MACHINE, oemKeyPath, "OEMKey", oemValue
If Err.Number = 0 And oemValue <> "" Then
    strLicenseType = "OEM"
Else
    strLicenseType = "Retail"
End If
On Error GoTo 0

' Get the currently logged-in username
Set colUsers = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objUser In colUsers
    strUserName = objUser.UserName
Next

' Output Message
strPopupMsg = strOS & vbNewLine & vbNewLine
strPopupMsg = strPopupMsg & "Version: " & strVersion & vbNewLine
strPopupMsg = strPopupMsg & "Build Number: " & strBuild & vbNewLine
strPopupMsg = strPopupMsg & "PID: " & strSerial & vbNewLine
strPopupMsg = strPopupMsg & "Registered to: " & strRegistered & vbNewLine
strPopupMsg = strPopupMsg & "Your Windows Product Key is: " & strProductKey & vbNewLine
strPopupMsg = strPopupMsg & "Manufacturer: " & strManufacturer & vbNewLine
strPopupMsg = strPopupMsg & "Model: " & strModel & vbNewLine
strPopupMsg = strPopupMsg & "License Type: " & strLicenseType & vbNewLine
strPopupMsg = strPopupMsg & "Currently Logged In User: " & strUserName

' Display the message
Set wshShell = CreateObject("wscript.shell")
wshShell.Popup strPopupMsg, 0, "Windows License Information", vbInformation
