Attribute VB_Name = "ModErrorHandling"
'===============================================================
' Module ModErrorHandling
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 20 Apr 18
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModErrorHandling"

Public FaultCount1002 As Integer
Public FaultCount1008 As Integer

' ===============================================================
' CentralErrorHandler
' Handles all system errors
' ---------------------------------------------------------------
Public Function CentralErrorHandler( _
            ByVal ErrModule As String, _
            ByVal ErrProc As String, _
            Optional ByVal ErrFile As String, _
            Optional ByVal EntryPoint As Boolean) As Boolean

    Static ErrMsg As String
    
'    Dim iFile As Integer
    Dim ErrNum As Long
    Dim ErrHeader As String
    Dim SySysPath As String
    Dim LogText As String
'    Dim ErrMsgTxt As String
    
    ErrNum = Err.Number
    
    If Len(ErrMsg) = 0 Then ErrMsg = Err.Description
                
    On Error Resume Next
    
    If Len(ErrFile) = 0 Then ErrFile = ThisWorkbook.Name
    
    SysPath = ThisWorkbook.Path
    
    If Right$(SysPath, 1) <> "\" Then SysPath = SysPath & "\"
    
    SysPath = SysPath & "System Files\"
    
    ErrHeader = "[" & ErrFile & "]" & ErrModule & "." & ErrProc

    LogText = "  " & ErrHeader & ", Error " & CStr(lErrNum) & ": " & sErrMsg
    
    If OUTPUT_MODE = "Log" Then
        Dim Response As Integer
        
        iFile = FreeFile()
        Open SysPath & FILE_ERROR_LOG For Append As #iFile
        Print #iFile, Format$(Now(), "mm/dd/yy hh:mm:ss"); sLogText
        If bEntryPoint Then Print #iFile,
        Close #iFile
    End If
                
    Debug.Print Format$(Now(), "mm/dd/yy hh:mm:ss"); sLogText
    If bEntryPoint Then Debug.Print
    
    If bEntryPoint Or DEBUG_MODE Then
        ModLibrary.PerfSettingsOff

        If Not ModLibrary.OutlookRunning Then
            Shell "Outlook.exe"
        End If

        If MailSystem Is Nothing Then Set MailSystem = New ClsMailSystem
    
        If Not DEV_MODE And SEND_ERR_MSG Then SendErrMessage
            
        sErrMsg = vbNullString
    End If
    
    CentralErrorHandler = DEBUG_MODE
    
End Function

' ===============================================================
' CustomErrorHandler
' Handles system custom errors 1000 - 1500
' ---------------------------------------------------------------
Public Function CustomErrorHandler(ErrorCode As Long, Optional Message As String) As Boolean
    Dim MailSubject As String
    Dim MailBody As String
    
    Const StrPROCEDURE As String = "CustomErrorHandler()"

    On Error Resume Next

    Select Case ErrorCode
        Case UNKNOWN_USER
            MailSubject = "Unknown User - " & APP_NAME
            MailBody = "A new user needs to be added to the database - " & CurrentUser.CrewNo & " " & CurrentUser.UserName
                                
            If Not ModReports.SendEmailReports(MailSubject, MailBody, EnumNewGuestUser) Then Err.Raise HANDLED_ERROR
            
            MsgBox "Sorry, the system does not recognise you.  Please continue with " _
                    & "the order as a guest.  Your name has been forwarded onto the " _
                    & "Administrator so that you can be added to the system", vbOKOnly + vbInformation, APP_NAME
                               
            CurrentUser.AddTempAccount
            
            CurrentUser.DBSave
            
        Case NO_ITEM_SELECTED
            MsgBox "Please select an item", vbOKOnly + vbInformation, APP_NAME
            
        Case NO_DATABASE_FOUND
            FaultCount1008 = FaultCount1008 + 1
            Debug.Print "Trying to connect to Database....Attempt " & FaultCount1008
            
            If ModErrorHandling.FaultCount1008 <= 3 Then
            
                Application.DisplayStatusBar = True
                Application.StatusBar = "Trying to connect to Database....Attempt " & FaultCount1008
                Application.Wait (Now + TimeValue("0:00:02"))
                Debug.Print FaultCount1008
            Else
                FaultCount1008 = 0
                Application.StatusBar = "System Failed - No Database"
                End
            End If
        
        Case SYSTEM_RESTART
            Debug.Print "system failed - restarting"
            FaultCount1002 = FaultCount1002 + 1

            If ModErrorHandling.FaultCount1002 <= 3 Then
                If Not Initialise Then Err.Raise HANDLED_ERROR
                Application.DisplayStatusBar = True
                Application.StatusBar = "System failed...Restarting Attempt " & FaultCount1002
                Application.Wait (Now + TimeValue("0:00:02"))
            Else
                FaultCount1002 = 0
                Application.StatusBar = "Sysetm Failed"
                End
            End If
            
        Case NO_QUANTITY_ENTERED
            MsgBox "Please enter a quantity", vbExclamation, APP_NAME
        
        Case NO_SIZE_ENTERED
            MsgBox "Please enter a size", vbExclamation, APP_NAME
        
        Case NO_CREW_NO_ENTERED
            MsgBox "Please enter a Brigade No", vbExclamation, APP_NAME
            
        Case NUMBERS_ONLY
            MsgBox "Please enter number only", vbExclamation, APP_NAME
            
        Case CREWNO_UNRECOGNISED
            MsgBox "The Brigade No is not recognised on the system, please re-enter", vbExclamation, APP_NAME
        
        Case NO_VEHICLE_SELECTED
            MsgBox "Please select a vehicle", vbExclamation, APP_NAME
        
        Case NO_STATION_SELECTED
            MsgBox "Please select a station", vbExclamation, APP_NAME
            
        Case FIELDS_INCOMPLETE
            MsgBox "Please complete all fields", vbExclamation, APP_NAME
            
        Case NO_NAMES_SELECTED
            MsgBox "Please select a name", vbExclamation, APP_NAME
            
        Case FORM_INPUT_EMPTY
            MsgBox "Please complete all highlighted fields", vbExclamation, APP_NAME
            
        Case ACCESS_DENIED
            MsgBox "Sorry you do not have the required Access Level.  " _
                & "Please send a Support Mail if you require access", vbCritical, APP_NAME
        
        Case NO_ORDER_MESSAGE
            MsgBox Message
            
        Case NO_INI_FILE
            MsgBox "No INI file has been found, so system cannot continue. This can occur if the file " _
                    & "is copied from its location on the T Drive.  Please delete file and create a shortcut instead", vbCritical, APP_NAME
            Application.StatusBar = "System Failed - No INI File"
            End
        
        Case NO_STOCK_AVAIL
            MsgBox "You cannot issue this item as there insuficient stock available", vbExclamation, APP_NAME
            
        Case DB_WRONG_VER
            MsgBox "Incorrect Version Database - System cannot continue", vbCritical + vbOKOnly, APP_NAME
            Application.StatusBar = "System Failed - Wrong DB Version"
            End
        
        Case NO_FILE_SELECTED
            MsgBox "There was no file selected", vbOKOnly + vbExclamation, APP_NAME
        Exit Sub

    End Select
    
    Set MailSystem = Nothing

    CustomErrorHandler = True


End Sub

Private Sub SendErrMessage()
    With MailSystem
        .MailItem.To = "Julian Turner"
        .MailItem.Subject = "Debug Report - " & APP_NAME
        .MailItem.Importance = olImportanceHigh
        .MailItem.Attachments.Add SysPath & FILE_ERROR_LOG
        .MailItem.Body = "Please add any further information such " _
                           & "what you were doing at the time of the error" _
                           & ", and what candidate were you working on etc "
                           
        If TEST_MODE Then .DisplayEmail Else .SendEmail
    End With

End Sub
