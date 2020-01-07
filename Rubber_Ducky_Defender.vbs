'File Name: Rubber_Ducky_Defender.vbs
'Version: v1.1, 1/6/2020
'Author: Justin Grimes, 1/3/2020

'-------------------------------------------------- 
'Specify which global variables will be used in this script.
Option Explicit
Dim strComputer, objWMIService, objNet, objFSO, colMonitoredEvents, objShell, wmiServices, query, return1, return2, objLatestEvent, _
 param1, param2, param3, param4, param5, usbOnly, arg, userName, hostName, mailFile, mFile, mailData, strComputerName, _
 resultCounter, strSafeDate, strSafeTime, strDateTime, strLogFilePath, strLogFileName, returnData, objLogFile, emailDisable, _
 logDisable, guiDisable, strSafeTimeRAW, strSafeTimeDIFF, strSafeTimeLAST, disableThreats, query2, objDevice, colDevices, _
 strDeviceName, strDeviceNames, colDevice, DevCaption, DevID, DevInstallDate, appPath, company, companyAbbreviation, detectionArray, _
 fromEmail, toEmail, sendmailPath, logPath, arrDeviceNames, colUSBDevice, colUSBDevices, detected, mArray, mValue, mValEl, appName, _
 warnOnThreat, confirmationBox, warnFlag, killFlag
'-------------------------------------------------- 

' ----------
' SET THESE VARIABLES TO YOUR ENVIRONMENT!!!
company = "Company Inc."
companyAbbreviation = "Company"
fromEmail = "Server@company.com"
toEmail = "IT@company.com"
sendmailPath = "sendmail.exe"
logPath = "\\server\Logs"
emailDisable = FALSE
logDisable = FALSE
guiDisable = FALSE
disableThreats = TRUE
warnOnThreat = TRUE
appPath = "\\server\scripts\Rubber_Ducky_Defender"
appName = "Rubber_Ducky_Defender"
' ---------- 

'-------------------------------------------------- 
'Define variables for the session.
strComputer = "." 
resultCounter = 0
param1 = ""
param2 = ""
strSafeTimeRAW = 0
strSafeTimeDIFF = 0
strSafeTimeLAST = 0
detected = FALSE
killFlag = FALSE
confirmationBox = FALSE
detectionArray = Array()
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colMonitoredEvents = objWMIService.ExecNotificationQuery("SELECT * FROM __InstanceCreationEvent WITHIN 10 WHERE Targetinstance " & _ 
 "ISA 'Win32_PNPEntity' and TargetInstance.DeviceId like '%HID%'") 
Set wmiServices = GetObject ("winmgmts:{impersonationLevel=Impersonate}!//" & strComputer)
Set arg = WScript.Arguments
Set objNet = CreateObject("Wscript.Network") 
Set objShell = WScript.CreateObject("WScript.Shell")
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
userName = objNet.Username 
hostName = objNet.Computername
mailFile = "C:\Users\" & userName & "\RD_Warning.mail"
'--------------------------------------------------

'--------------------------------------------------
'Retrieve the specified arguments.
If (arg.Count > 0) Then
  param1 = arg(0)
End If
If (arg.Count > 1) Then
  param2 = arg(1)
End If
If (arg.Count > 2) Then
  param3 = arg(2)
End If
If (arg.Count > 3) Then
  param4 = arg(3)
End If
If (arg.Count > 4) Then
  param5 = arg(4)
End If
'If the -e or --email arguments are set we disable the notification email.
If (param1 = "-e" Or param1 = "--email") Then
  emailDisable = TRUE
End If
If (param2 = "-e" Or param2 = "--email") Then
  emailDisable = TRUE
End If
If (param3 = "-e" Or param3 = "--email") Then
  emailDisable = TRUE
End If
If (param4 = "-e" Or param4 = "--email") Then
  emailDisable = TRUE
End If
If (param5 = "-e" Or param5 = "--email") Then
  emailDisable = TRUE
End If
'If the -l or --log arguments are set we disable the logfile.
If (param1 = "-l" Or param1 = "--log") Then
  logDisable = TRUE
End If
If (param2 = "-l" Or param2 = "--log") Then
  logDisable = TRUE
End If
If (param3 = "-l" Or param3 = "--log") Then
  logDisable = TRUE
End If
If (param4 = "-l" Or param4 = "--log") Then
  logDisable = TRUE
End If
If (param5 = "-l" Or param5 = "--log") Then
  logDisable = TRUE
End If
'If the -g or --gui arguments are set we disable the GUI.
If (param1 = "-g" Or param1 = "--gui") Then
  guiDisable = TRUE
End If
If (param2 = "-g" Or param2 = "--gui") Then
  guiDisable = TRUE
End If
If (param3 = "-g" Or param3 = "--gui") Then
  guiDisable = TRUE
End If
If (param4 = "-g" Or param4 = "--gui") Then
  guiDisable = TRUE
End If
If (param5 = "-g" Or param4 = "--gui") Then
  guiDisable = TRUE
End If
'--------------------------------------------------

'--------------------------------------------------
'A funciton for running SendMail.
Function SendEmail() 
  objShell.exec "C:\Windows\System32\cmd.exe /c " & sendmailPath & " " & mailFile
End Function
'--------------------------------------------------

'--------------------------------------------------
'Perform the loop that checks for new devices.
Do While TRUE
  Set objLatestEvent = colMonitoredEvents.NextEvent 
  If (resultCounter = 0) Then
    query = "Select * From Win32_USBControllerDevice"
    Set colDevices = objWMIService.ExecQuery(query)
    'Loop through the list of returned devices to gain more information about what was connected.
    For Each objDevice In colDevices
      strDeviceName = Replace(objDevice.Dependent, Chr(34), "")
      arrDeviceNames = Split(strDeviceName, "=")
      strDeviceName = arrDeviceNames(1)
      If InStr(" " & strDeviceName, "HID") Then
        query2 = "Select * From Win32_PnPEntity Where DeviceID = '" & strDeviceName & "'"
        Set colUSBDevices = objWMIService.ExecQuery(query2)
        query2 = ""
        return1 = ""
        'Build the return data.
        For Each colUSBDevice In colUSBDevices
          DevCaption = colUSBDevice.Caption
          DevID = colUSBDevice.DeviceID
          DevInstallDate = colUSBDevice.InstallDate
          If (DevCaption = "HID Keyboard Device") Then
            resultCounter = resultCounter + 1
            return1 = "Device Type: " & DevCaption & ", " & _
            vbNewLine & "Device ID: " & DevID & ", " & _
            vbNewLine & "Date Installed: " & colUSBDevice.InstallDate & _
            vbNewLine & vbNewLine
            return2 = return1 & return2
            detected = TRUE
            If (warnOnThreat = TRUE) Then
              warnFlag = TRUE
            End If
            If (disableThreats = TRUE) Then
              killFlag = TRUE
            End If
          End If
        Next 
      End If
    Next
  End IF
  strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
  strSafeTimeRAW = strSafeTime
  strSafeTimeDIFF = strSafeTime - strSafeTimeLAST
  returnData = Notify()
  If (strSafeTimeDIFF > 30) Then
    If (warnFlag = TRUE) Then
      confirmationBox = MsgBox("The device you just plugged in reports that it is a USB Keyboard. Did you intend to plug in a keyboard?", 4, appName)
    Else
      confirmationBox = 7
    End If
    If (disableThreats = TRUE And killFlag = TRUE And confirmationBox = 7) Then
      If (logDisable = FALSE) Then 
        CreateLog(returnData)
      End If
      warnFlag = FALSE
      killFlag = FALSE
      killWorkstation()
    End If
  End If 
  strSafeTimeLAST = strSafeTimeRAW
Loop
'--------------------------------------------------

'--------------------------------------------------
'A function to format the notification email and notify the user.
function Notify()
  If (resultCounter > 0) Then
    resultCounter = resultCounter - 1
  End If
  If (resultCounter = 0 And detected = TRUE) Then
    'Prepare the notification email and popup.
    Set mFile = objFSO.CreateTextFile(mailFile, TRUE, FALSE)  
    mFile.Write "To: " & toEmail & vbNewLine & "From: " & fromEmail & vbNewLine & "Subject: " & companyAbbreviation & " New USB Input Device Connected!!!" & _
     vbNewLine & "This is an automatic email from the " & company & " Network to notify you that a potentially dangerous device was detected on a domain workstation." & _
     vbNewLine & vbNewLine & _
     "Please review the information below to verify that the connected device is not a threat." & _
     vbNewLine & vbNewLine & _
     "DEVICE DETAILS: " & _
     vbNewLine & vbNewLine & _
     "Workstation: " & hostName & ", " & _
     vbNewLine & "Username: " & userName & ", " & _
     vbNewLine & vbNewLine & "Detected Devices: " & _
     vbNewLine &vbNewLine & return2 & vbNewLine & _
     "This check was generated by " & strComputerName & " and is run in the background upon user logon." & _
     vbNewLine & vbNewLine & _
     "Script: ""Rubber_Ducky_Defender.vbs""" 
    mFile.Close
    If (emailDisable = FALSE) Then
      SendEmail
    End If
    'Display results.
    If (guiDisable = FALSE And detected = TRUE) Then
      mailData = "Devices Detected: " & vbNewLine & vbNewLine & return2
      MsgBox mailData, vbOKOnly, appName
    End If
    'Reset the outputs for the next iteration of the loop above. (MUST BE DONE!!! This was the source of a lot of debugging.)
    detected = FALSE
    Notify = return2
    return2 = ""
    return1 = ""
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to create a log file.
Function CreateLog(strEventInfo)
  If Not (strEventInfo = "") Then
    'Logfile related variables are defined at log creation time for accurate time reporting.
    strLogFilePath = logPath
    strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
    strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
    strSafeTimeRAW = strSafeTime
    strSafeTimeDIFF = strSafeTime - strSafeTimeLAST
    strLogFileName = strLogFilePath & "\" & userName & "-" & strDateTime & "-rubber_ducky_defender.txt"
    Set objLogFile = objFSO.CreateTextFile(strLogFileName, TRUE, FALSE)
    objLogFile.WriteLine(strEventInfo)
    objLogFile.Close
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function shut down the machine when triggered.
Function killWorkstation()
  'oShell.Run "C:\Windows\System32\shutdown.exe /s /f /t 0", 0, FALSE
  Msgbox "Uncomment the oShell.Run line in Rubber_Ducky_Defender.vbs to enable automatic shutdown upon detection!"
End Function
'--------------------------------------------------