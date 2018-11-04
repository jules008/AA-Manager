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
        Debug.Print SQL
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
        
    Debug.Print "Connect to DB: " & DB_PATH & DB_FILE_NAME
    
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
    
    If DB Is Nothing Then
        ReadINIFile
        DBConnect
    End If
    
    DB.Execute "CREATE TABLE TblDBVersion"
    DB.Execute "ALTER TABLE TblDBVersion ADD Version Text"
    
    Set RstTable = SQLQuery("TblDBVersion")

    With RstTable
        .AddNew
        .Fields(0) = "V0.0.0"
        .Update
    End With
       
    'check preceding DB Version
    If RstTable!VERSION <> "V0.0.0" Then
        MsgBox "Database needs to be upgraded to V0.0.0 to continue", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    'Table TblContractLookup
    DB.Execute "CREATE TABLE TblContractLookup"
    DB.Execute "ALTER TABLE TblContractLookup ADD ContractNo Long"
    DB.Execute "ALTER TABLE TblContractLookup ADD ContractType Text"
    DB.Execute "INSERT INTO TblContractLookup VALUES (1, 'Under 120 Hrs')"
    DB.Execute "INSERT INTO TblContractLookup VALUES (2, 'Over 120 Hrs')"

    'Table TblStnLookUp
    DB.Execute "CREATE TABLE TblStnLookUp"
    DB.Execute "ALTER TABLE TblStnLookUp ADD StationNo Long"
    DB.Execute "ALTER TABLE TblStnLookUp ADD Callsign Text"
    DB.Execute "ALTER TABLE TblStnLookUp ADD Name Text"
    DB.Execute "ALTER TABLE TblStnLookUp ADD Address Text"
    DB.Execute "ALTER TABLE TblStnLookUp ADD StationType Text"
    DB.Execute "ALTER TABLE TblStnLookUp ADD Division Text"
    
    DB.Execute "INSERT INTO TblStnLookUp VALUES (1 , 'EC01', 'Alford', 'Willoughby Rd, Alford LN13 9AT, UK', 2, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (2 , 'EC02', 'Bardney', 'Bardney, Lincoln LN3 5TF, UK', 2, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (3 , 'EC03', 'Billingborough', 'High St, Billingborough, Sleaford NG34 0QA, UK', 2, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (4 , 'EC04', 'Billinghay', 'Mill Ln, Billinghay, Lincoln LN4 4ES, UK', 2, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (5 , 'EC05', 'Binbrook', 'St Marys Ln, Binbrook, Market Rasen LN8 6DL, UK', 2, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (6 , 'EC06', 'Boston', 'Robin Hoods Walk, Boston PE21 9EP, UK', 1, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (7 , 'EC07', 'Bourne', 'South St, Bourne PE10 9LY, UK', 2, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (8 , 'EC08', 'Brant Broughton', 'High St, Brant Broughton, Lincoln LN5 0SL, UK', 2, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (9 , 'EC09', 'Caistor', 'Caistor, Market Rasen LN7, UK', 2, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (10 , 'EC10', 'Corby Glen', 'Bourne Rd, United Kingdom', 2, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (11 , 'EC11', 'Crowland', 'Thorney Rd, Crowland, Peterborough PE6 0AL, UK', 2, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (12 , 'EC12', 'Donington', 'High St, Donington, Spalding PE11 4TA, UK', 2, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (13 , 'EC13', 'Gainsborough', 'Nelson St, Gainsborough DN21 2SE, UK', 1, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (14 , 'EC14', 'Grantham', 'Fire Station/Harlaxton Rd, Grantham NG31 7SG, UK', 1, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (15 , 'EC15', 'Holbeach', 'Holbeach, Spalding PE12, UK', 2, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (16 , 'EC16', 'Horncastle', 'Foundry St, Horncastle LN9 6AB, UK', 2, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (17 , 'EC17', 'Kirton', 'Station Rd, Kirton, Boston PE20 1LD, UK', 2, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (18 , 'EC18', 'Leverton', 'Old Main Rd, Old Leake, Boston PE22 9HT, UK', 2, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (19 , 'EC19', 'Lincoln North', 'Nettleham Road, Lincoln, LN2 4DH.', 1, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (20 , 'EC20', 'Lincoln South', 'South Pk Av, Lincoln LN5 8EL, UK', 1, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (21 , 'EC21', 'Long Sutton', 'Long Sutton, Spalding PE12, UK', 2, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (22 , 'EC22', 'Louth', 'Eastfield Rd, Louth LN11 7AS, UK', 1, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (23 , 'EC23', 'Mablethorpe', 'The Blvd, Mablethorpe LN12 2AD, UK', 2, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (24 , 'EC24', 'Market Deeping', 'High Street Market, High St, Market Deeping, Peterborough PE6 8ED, UK', 2, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (25 , 'EC25', 'Market Rasen', 'Linwood Rd, Market Rasen LN8 3AN, UK', 2, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (26 , 'EC26', 'Metheringham', 'Fen Road, Metheringham, LN4 3AA.', 2, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (27 , 'EC27', 'North Hykeham', 'Mill Ln, North Hykeham, Lincoln LN6 9PE, UK', 2, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (28 , 'EC28', 'North Somercotes', 'Churchill Rd, North Somercotes, Louth LN11, UK', 2, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (29 , 'EC29', 'Saxilby', 'Saxilby, Lincoln LN1, UK', 2, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (30 , 'EC30', 'Skegness', 'Churchill Ave, Skegness PE25 2RN, UK', 1, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (31 , 'EC31', 'Sleaford', 'Church Ln, Sleaford NG34, UK', 2, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (32 , 'EC32', 'Spalding', 'High St, Donington, Spalding PE11 4TA, UK', 2, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (33 , 'EC33', 'Spilsby', 'Boston Rd, Spilsby PE23 5HH, UK', 2, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (34 , 'EC34', 'Stamford', '68 New Cross Rd, Stamford PE9 1, UK', 2, 'South')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (35 , 'EC35', 'Waddington', 'Mere Road, Waddington, LN5 9NX.', 2, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (36 , 'EC36', 'Wainfleet', 'Magdalen Rd, Wainfleet All Saints, Skegness PE24 4DD, UK', 2, 'East')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (37 , 'EC37', 'Woodhall Spa', 'King Edward Rd, Woodhall Spa LN10 6RL, UK', 2, 'West')"
    DB.Execute "INSERT INTO TblStnLookUp VALUES (38 , 'EC38', 'Wragby', 'Millbrook Ln, Wragby, Market Rasen LN8 5AB, UK', 2, 'West')"

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
    DB.Execute "ALTER TABLE TblTemplate DROP ID, NoStation, StationNo, StationName"
    DB.Execute "ALTER TABLE TblTemplate ADD ContractType Double, HrsPW Double, NoWeeks Double, RevDateDue Date"
    DB.Execute "ALTER TABLE TblTemplate ALTER COLUMN Role Long"

    'Table TblTemplateStns
    DB.Execute "SELECT * INTO TblTemplateStns FROM TblTemplateBAK"
    DB.Execute "ALTER TABLE TblTemplateStns DROP ID"
    DB.Execute "ALTER TABLE TblTemplateStns DROP Role"
    DB.Execute "ALTER TABLE TblTemplateStns DROP CrewName"
    DB.Execute "ALTER TABLE TblTemplateStns DROP StationName"
    DB.Execute "ALTER TABLE TblTemplateStns DROP TemplateDate"
    DB.Execute "ALTER TABLE TblTemplateStns ADD HrsPW Double"
    
    Dim tbl As TableDef
    Set tbl = DB.TableDefs("TblTemplateStns")
    tbl.Fields("NoStation").Name = "Station"
    
    'Table TemplateDetail
    DB.Execute "SELECT * INTO TblTemplateDetail FROM TemplateDetail"
    DB.Execute "SELECT * INTO TblTemplateDetailBAK FROM TemplateDetail"
    DB.Execute "DROP TABLE TemplateDetail"
    DB.Execute "ALTER TABLE TblTemplateDetail DROP ID1, StationNo, ClosedDate"
    DB.Execute "ALTER TABLE TblTemplateDetail ALTER COLUMN OnCall Double"
    
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
                    & " VALUES ('5398', 'Julian', 'Turner', 'Julian Turner', 'Admin', TRUE, 2, TRUE, '1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1')"
    
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
    Set tbl = Nothing
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
        
    If DB Is Nothing Then
        ReadINIFile
        DBConnect
    End If
    
    Set RstTable = SQLQuery("TblDBVersion")

    If RstTable.Fields(0) <> "V0.0.1" Then
        MsgBox "Database needs to be upgraded to V0.0.1 to continue", vbOKOnly + vbCritical
        Exit Sub
    End If
       
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
    DB.Execute "DROP TABLE TblTemplateStns"
    
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
