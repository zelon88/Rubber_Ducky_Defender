'File Name: Rubber_Ducky_Defender.vbs
'Version: v1.3, 1/8/2020
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

'The complete, unabbreviated name of your organization.
company = "Company Inc."
'The abbreviated name of your organization.
companyAbbreviation = "Company"
'The email address where notification emails will appear to have come from.
fromEmail = "Server@company.com"
'The email address where notification emails should be sent to.
toEmail = "IT@company.com"
'The full, absolute UNC path to the location where sendmail.exe is located.
sendmailPath = "sendmail.exe"
'The full, absolute UNC path to the location where logs can be stored.
logPath = "\\server\Logs"
'Set to TRUE to disable emails by default, regardless of command line arguments.
emailDisable = FALSE
'Set to TRUE to disable logging by default, regardless of command line arguments.
logDisable = FALSE
'Set to TRUE to disable the user interface (message boxes), regardless of command line arguments.
guiDisable = FALSE
'Set to TRUE to mitigate detections that are found. If set to FALSE only notifications will take place.
disableThreats = TRUE
'Set to true to fire a confirmation box upon detection asking the user to confirm their intention to connect a USB keyboard.
warnOnThreat = TRUE
'The full, absolute UNC path to the directory where this script is stored.
appPath = "\\server\scripts\Rubber_Ducky_Defender"
'The file name of this script. 
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
'If the -ne or --nemail arguments are set we disable the notification email.
If (param1 = "-ne" Or param1 = "--nemail") Then
  emailDisable = TRUE
End If
If (param2 = "-ne" Or param2 = "--nemail") Then
  emailDisable = TRUE
End If
If (param3 = "-ne" Or param3 = "--nemail") Then
  emailDisable = TRUE
End If
If (param4 = "-ne" Or param4 = "--nemail") Then
  emailDisable = TRUE
End If
If (param5 = "-ne" Or param5 = "--nemail") Then
  emailDisable = TRUE
End If
'If the -nl or --nlog arguments are set we disable the logfile.
If (param1 = "-nl" Or param1 = "--nlog") Then
  logDisable = TRUE
End If
If (param2 = "-nl" Or param2 = "--nlog") Then
  logDisable = TRUE
End If
If (param3 = "-nl" Or param3 = "--nlog") Then
  logDisable = TRUE
End If
If (param4 = "-nl" Or param4 = "--nlog") Then
  logDisable = TRUE
End If
If (param5 = "-nl" Or param5 = "--nlog") Then
  logDisable = TRUE
End If
'If the -ng or --ngui arguments are set we disable the GUI.
If (param1 = "-ng" Or param1 = "--ngui") Then
  guiDisable = TRUE
End If
If (param2 = "-ng" Or param2 = "--ngui") Then
  guiDisable = TRUE
End If
If (param3 = "-ng" Or param3 = "--ngui") Then
  guiDisable = TRUE
End If
If (param4 = "-ng" Or param4 = "--ngui") Then
  guiDisable = TRUE
End If
If (param5 = "-ng" Or param4 = "--ngui") Then
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
    'Loop through the list of available USB controllers.
    For Each objDevice In colDevices
      strDeviceName = Replace(objDevice.Dependent, Chr(34), "")
      arrDeviceNames = Split(strDeviceName, "=")
      strDeviceName = arrDeviceNames(1)
      'Drill into each USB controller and enumerate the devices attached to it.
      If InStr(" " & strDeviceName, "HID") Then
        query2 = "Select * From Win32_PnPEntity Where DeviceID = '" & strDeviceName & "'"
        Set colUSBDevices = objWMIService.ExecQuery(query2)
        query2 = ""
        return1 = ""
        'Look for USB keyboards among the enumerated devices.
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
            'Upon detection, set the "warnFlag" to fire the user confirmation box (if enabled).
            If (warnOnThreat = TRUE) Then
              warnFlag = TRUE
            End If
            'Upon detection, set the "killFlag" to fire the killWorkstation() function (if enabled). 
            If (disableThreats = TRUE) Then
              killFlag = TRUE
            End If
          End If
        Next 
      End If
    Next
  End IF
  'Reset time variables detected on the last loop. Used to prevent erroneous double-detections.
  strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
  strSafeTimeRAW = strSafeTime
  strSafeTimeDIFF = strSafeTime - strSafeTimeLAST
  returnData = Notify()
  If (strSafeTimeDIFF > 30) Then
    If (warnFlag = TRUE And guiDisable = FALSE) Then
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
     "Script: """ & appName & ".vbs""" 
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
    'Logfile related time variables are defined at log creation time for accurate time reporting.
    strLogFilePath = logPath
    strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
    strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
    strSafeTimeRAW = strSafeTime
    strSafeTimeDIFF = strSafeTime - strSafeTimeLAST
    strLogFileName = strLogFilePath & "\" & userName & "-" & strDateTime & "-" & appName & ".txt"
    Set objLogFile = objFSO.CreateTextFile(strLogFileName, TRUE, FALSE)
    objLogFile.WriteLine(strEventInfo)
    objLogFile.Close
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function shut down the machine when triggered.
Function killWorkstation()
  'objShell.Run "C:\Windows\System32\shutdown.exe /s /f /t 0", 0, FALSE
  Msgbox "Uncomment the objShell.Run line in " & appName & ".vbs to enable automatic shutdown upon detection!", vbOKOnly, appName
End Function
'--------------------------------------------------