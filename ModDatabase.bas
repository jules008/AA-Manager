Attribute VB_Name = "ModDatabase"
'===============================================================
' Module ModDatabase
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Apr 18
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModDatabase"

Public DB As DAO.Database
Public MyQueryDef As DAO.QueryDef

' ===============================================================
' SQLQuery
' Queries database with given SQL script
' ---------------------------------------------------------------
Public Function SQLQuery(SQL As String) As Recordset
    Dim RstResults As Recordset
    
    Const StrPROCEDURE As String = "SQLQuery()"

    On Error GoTo ErrorHandler
      
Restart:
    Application.StatusBar = ""

    If DB Is Nothing Then
        Err.Raise NO_DATABASE_FOUND, Description:="Unable to connect to database"
    Else
        If FaultCount1008 > 0 Then FaultCount1008 = 0
    
        Set RstResults = DB.OpenRecordset(SQL, dbOpenDynaset)
        Set SQLQuery = RstResults
    End If
    
    Set RstResults = Nothing
Exit Function

ErrorExit:

    Set RstResults = Nothing

    Set SQLQuery = Nothing
    Terminate

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        If CustomErrorHandler(Err.Number) Then
            If Not Initialise Then Err.Raise HANDLED_ERROR
            Resume Restart
        Else
            Err.Raise HANDLED_ERROR
        End If
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' DBConnect
' Provides path to database
' ---------------------------------------------------------------
Public Function DBConnect() As Boolean
    Const StrPROCEDURE As String = "DBConnect()"

    On Error GoTo ErrorHandler

    Set DB = OpenDatabase(DB_PATH & DB_FILE_NAME)
  
    DBConnect = True

Exit Function

ErrorExit:

    DBConnect = False

Exit Function

ErrorHandler:

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' DBTerminate
' Disconnects and closes down DB connection
' ---------------------------------------------------------------
Public Function DBTerminate() As Boolean
    Const StrPROCEDURE As String = "DBTerminate()"

    On Error GoTo ErrorHandler

    If Not DB Is Nothing Then DB.Close
    Set DB = Nothing

    DBTerminate = True

Exit Function

ErrorExit:

    DBTerminate = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' SelectDB
' Selects DB to connect to
' ---------------------------------------------------------------
Public Function SelectDB() As Boolean
    Const StrPROCEDURE As String = "SelectDB()"

    On Error GoTo ErrorHandler
    Dim DlgOpen As FileDialog
    Dim FileLoc As String
    Dim NoFiles As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'open files
    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    
     With DlgOpen
        .Filters.Clear
        .Filters.Add "Access Files (*.accdb)", "*.accdb"
        .AllowMultiSelect = False
        .Title = "Connect to Database"
        .Show
    End With
    
    'get no files selected
    NoFiles = DlgOpen.SelectedItems.Count
    
    'exit if no files selected
    If NoFiles = 0 Then
        MsgBox "There was no database selected", vbOKOnly + vbExclamation, "No Files"
        SelectDB = True
        Exit Function
    End If
  
    'add files to array
    For i = 1 To NoFiles
        FileLoc = DlgOpen.SelectedItems(i)
    Next
    
    DB_PATH = FileLoc
    
    Set DlgOpen = Nothing

    SelectDB = True

Exit Function

ErrorExit:

    Set DlgOpen = Nothing
    SelectDB = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UpdateDBScript
' Script to update DB
' ---------------------------------------------------------------
Public Sub UpdateDBScript()
    Dim TableDef As DAO.TableDef
    Dim Ind As DAO.Index
    Dim RstTable As Recordset
    Dim i As Integer
    Dim Binary As String
    
    Dim Fld As DAO.Field
    
    DBConnect
    
    Set DB = OpenDatabase("\\lincsfire.lincolnshire.gov.uk\folderredir$\Documents\julian.turner\Documents\RDS Project\AA Manager\Dev Environment\System Files\Rappel Data Pre-Live v0,04.accdb")

    DB.Execute "CREATE TABLE TblDBVersion"
    DB.Execute "ALTER TABLE TblDBVersion ADD Version Text"
    
    Set RstTable = SQLQuery("TblDBVersion")

    With RstTable
        .AddNew
        .Fields(0) = "V0.0.0"
        .Update
    End With
    
    Set RstTable = SQLQuery("TblDBVersion")
    
    'check preceding DB Version
    If RstTable!VERSION <> "V0.0.0" Then
        MsgBox "Database needs to be upgraded to V0.0.0 to continue", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    'Table TblContractLookup
    DB.Execute "CREATE TABLE TblContractLookup"
    DB.Execute "ALTER TABLE TblContractLookup ADD ContractNo Long"
    DB.Execute "ALTER TABLE TblContractLookup ADD ContractType Text"
    
    'Table TblStnLookUp
    DB.Execute "CREATE TABLE TblStnLookUp"
    DB.Execute "ALTER TABLE TblStnLookUp ADD StationNo Long"
    DB.Execute "ALTER TABLE TblStnLookUp ADD Callsign Text"
    DB.Execute "ALTER TABLE TblStnLookUp ADD Name Text"
    DB.Execute "ALTER TABLE TblStnLookUp ADD Address Text"
    DB.Execute "ALTER TABLE TblStnLookUp ADD StationType Text"
    DB.Execute "ALTER TABLE TblStnLookUp ADD Division Text"
        
    'Table CrewMemberDetail
    DB.Execute "SELECT * INTO TblCrewMemberDetail FROM CrewMemberDetail"
    DB.Execute "DROP TABLE CrewMemberDetail"
    
    'Table CrewMember
    DB.Execute "SELECT * INTO TblCrewMember FROM CrewMember"
    DB.Execute "DROP TABLE CrewMember"
    
    'Table Station
    DB.Execute "SELECT * INTO TblStation FROM Station"
    DB.Execute "DROP TABLE Station"
    
    'Table StationDetail
    DB.Execute "SELECT * INTO TblStationDetail FROM StationDetail"
    DB.Execute "DROP TABLE StationDetail"
    
    'Table Template
    DB.Execute "SELECT * INTO TblTemplate FROM Template"
    DB.Execute "SELECT * INTO TblTemplateBAK FROM Template"
    DB.Execute "DROP TABLE Template"
    
    
    'Table TemplateDetail
    DB.Execute "SELECT * INTO TblTemplateDetail FROM TemplateDetail"
    DB.Execute "SELECT * INTO TblTemplateDetailBAK FROM TemplateDetail"
    DB.Execute "DROP TABLE TemplateDetail"
    
    'Table TimeTbl
    DB.Execute "SELECT * INTO TblTimeTbl FROM TimeTbl"
    DB.Execute "DROP TABLE TimeTbl"
    
    'Table TblPerson
    DB.Execute "CREATE TABLE TblPerson"
    DB.Execute "ALTER TABLE TblPerson ADD CrewNo Text"
    DB.Execute "ALTER TABLE TblPerson ADD Forename Text"
    DB.Execute "ALTER TABLE TblPerson ADD Surname Text"
    DB.Execute "ALTER TABLE TblPerson ADD Username Text"
    DB.Execute "ALTER TABLE TblPerson ADD RankGrade Text"
    DB.Execute "ALTER TABLE TblPerson ADD MailAlert yesno"
    DB.Execute "ALTER TABLE TblPerson ADD Role Long"
    DB.Execute "ALTER TABLE TblPerson ADD MessageRead YesNo"
    DB.Execute "ALTER TABLE TblPerson ADD Stations Text"


    DB.Execute "INSERT INTO TblPerson (Crewno,Forename,Surname,UserName,RankGrade,MailAlert, Role,MessageRead,Stations )" _
                    & " VALUES ('5398', 'Julian', 'Turner', 'Julian Turner', 'Admin', TRUE, 2, TRUE, 1)"
    
    'update DB Version
    Set RstTable = SQLQuery("TblDBVersion")
    
    With RstTable
        .Edit
        .Fields(0) = "V0.0.1"
        .Update
    End With
    
'    UpdateSysMsg
    
    MsgBox "Database successfully updated", vbOKOnly + vbInformation
    
    Set RstTable = Nothing
    Set TableDef = Nothing
    Set Fld = Nothing
    
End Sub
              
' ===============================================================
' UpdateDBScriptUndo
' Script to update DB
' ---------------------------------------------------------------
Public Sub UpdateDBScriptUndo()
    Dim TableDef As DAO.TableDef
    Dim Ind As DAO.Index
    Dim RstTable As Recordset
    Dim i As Integer
        
    Dim Fld As DAO.Field
        
    Set DB = OpenDatabase("\\lincsfire.lincolnshire.gov.uk\folderredir$\Documents\julian.turner\Documents\RDS Project\AA Manager\Dev Environment\System Files\Rappel Data Pre-Live v0,04.accdb")
    
    Set RstTable = SQLQuery("TblDBVersion")

    If RstTable.Fields(0) <> "V0.0.1" Then
        MsgBox "Database needs to be upgraded to V0.0.1 to continue", vbOKOnly + vbCritical
        Exit Sub
    End If
       
     
    Set RstTable = SQLQuery("TblDBVersion")

    With RstTable
        .Edit
        .Fields(0) = "V0.0.0"
        .Update
    End With
    
    Set RstTable = Nothing
    
    'Drop Table
    DB.Execute "DROP TABLE tblDBVersion"
    DB.Execute "DROP TABLE TblContractLookup"
    DB.Execute "DROP TABLE TblPerson"
    DB.Execute "DROP TABLE TblStnLookUp"
    
    'Table CrewMemberDetail
    DB.Execute "SELECT * INTO CrewMemberDetail FROM TblCrewMemberDetail"
    DB.Execute "DROP TABLE TblCrewMemberDetail"
    
    'Table CrewMember
    DB.Execute "SELECT * INTO CrewMember FROM TblCrewMember"
    DB.Execute "DROP TABLE TblCrewMember"
    
     'Table Station
    DB.Execute "SELECT * INTO Station FROM TblStation"
    DB.Execute "DROP TABLE TblStation"
    
     'Table StationDetail
    DB.Execute "SELECT * INTO StationDetail FROM TblStationDetail"
    DB.Execute "DROP TABLE TblStationDetail"
    
     'Table Template
    DB.Execute "SELECT * INTO Template FROM TblTemplateBAK"
    DB.Execute "DROP TABLE TblTemplate"
    DB.Execute "DROP TABLE TblTemplateBAK"
    
     'Table TemplateDetail
    DB.Execute "SELECT * INTO TemplateDetail FROM TblTemplateDetailBAK"
    DB.Execute "DROP TABLE TblTemplateDetail"
    DB.Execute "DROP TABLE TblTemplateDetailBAK"
    
     'Table TimeTbl
    DB.Execute "SELECT * INTO TimeTbl FROM TblTimeTbl"
    DB.Execute "DROP TABLE TblTimeTbl"
   
   MsgBox "Database reset successfully", vbOKOnly + vbInformation
    
    Set RstTable = Nothing
    Set TableDef = Nothing
    Set Fld = Nothing

End Sub

' ===============================================================
' GetDBVer
' Returns the version of the DB
' ---------------------------------------------------------------
Public Function GetDBVer() As String
    Dim DBVer As Recordset
    
    Const StrPROCEDURE As String = "GetDBVer()"

    On Error GoTo ErrorHandler

    Set DBVer = SQLQuery("TblDBVersion")

    GetDBVer = DBVer.Fields(0)

    Debug.Print DBVer.Fields(0)
    Set DBVer = Nothing
Exit Function

ErrorExit:

    GetDBVer = ""
    
    Set DBVer = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UpdateSysMsg
' Updates the system message and resets read flags
' ---------------------------------------------------------------
Public Sub UpdateSysMsg()
    Dim RstMessage As Recordset
    
    Set RstMessage = SQLQuery("TblMessage")
    
    With RstMessage
        If .RecordCount = 0 Then
            .AddNew
        Else
            .Edit
        End If
        
        .Fields("SystemMessage") = "Version " & VERSION & " - What's New" _
                    & Chr(13) & "(See Release Notes on Support tab for further information)" _
                    & Chr(13) & "" _
                    & Chr(13) & " - Bug Fix - Hidden Assets" _
                    & Chr(13) & ""
        
        .Fields("ReleaseNotes") = "Software Version: " & VERSION _
                    & Chr(13) & "Database Version: " & DB_VER _
                    & Chr(13) & "Date: " & VER_DATE _
                    & Chr(13) & "" _
                    & Chr(13) & "- Bug Fix - Hidden Assets - Had ANOTHER go at fixing the hidden assets bug.  Hopefully fixed now" _
                    & Chr(13) & ""
        .Update
    End With
    
    'reset read flags
    DB.Execute "UPDATE TblPerson SET MessageRead = False WHERE MessageRead = True"
    
    Set RstMessage = Nothing

End Sub

' ===============================================================
' ShowUsers
' Show users logged onto system
' ---------------------------------------------------------------
Public Sub ShowUsers()
    Dim RstUsers As Recordset
    
    Set RstUsers = SQLQuery("TblUsers")
    
    With RstUsers
        Debug.Print
        Do While Not .EOF
            Debug.Print "User: " & .Fields(0) & " - Logged on: " & .Fields(1)
            .MoveNext
        Loop
    End With
    
    Set RstUsers = Nothing
End Sub
