VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsPerson
' v0,5 - Initial Version
'---------------------------------------------------------------
' Date - 15 May 18
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
' GetUserName - Gets Username from CrewNo
' AddTempAccount - adds a temporary account in the database if user not recognised
'===============================================================

Option Explicit
Private pCrewNo As String
Private pForename As String
Private pSurname As String
Private pUsername As String
Private pRankGrade As String
Private pMailAlert As Boolean
Private pStations As String
Private pRole As EnumRole
Private pMessageRead As Boolean

'---------------------------------------------------------------
Public Property Get CrewNo() As String
    CrewNo = pCrewNo
End Property

Public Property Let CrewNo(ByVal vNewValue As String)
    pCrewNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Forename() As String
    Forename = pForename
End Property

Public Property Let Forename(ByVal vNewValue As String)
    pForename = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Surname() As String
    Surname = pSurname
End Property

Public Property Let Surname(ByVal vNewValue As String)
    pSurname = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get UserName() As String
    UserName = pUsername
End Property

Public Property Let UserName(ByVal vNewValue As String)
    pUsername = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get RankGrade() As String
    RankGrade = pRankGrade
End Property

Public Property Let RankGrade(ByVal vNewValue As String)
    pRankGrade = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get MailAlert() As Boolean
    MailAlert = pMailAlert
End Property

Public Property Let MailAlert(ByVal vNewValue As Boolean)
    pMailAlert = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Role() As String
    Role = pRole
End Property

Public Property Let Role(ByVal vNewValue As String)
    pRole = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get MessageRead() As Boolean
   MessageRead = pMessageRead
End Property

Public Property Let MessageRead(ByVal vNewValue As Boolean)
    pMessageRead = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Stations() As String
   Stations = pStations
End Property

Public Property Let Stations(ByVal vNewValue As String)
    pStations = vNewValue
End Property

' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet(UserNameOrCrewNo As String)
    Dim RstPerson As Recordset
    Dim UserName As String

    If Not IsNumeric(UserNameOrCrewNo) Then
        Set RstPerson = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE Username = '" & UserNameOrCrewNo & "'")
    Else
        Set RstPerson = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE CrewNo = '" & UserNameOrCrewNo & "'")
    End If
    
    With RstPerson
        If .RecordCount > 0 Then
            If Not IsNull(!CrewNo) Then pCrewNo = !CrewNo
            If Not IsNull(!Forename) Then pForename = !Forename
            If Not IsNull(!Surname) Then pSurname = !Surname
            If Not IsNull(!UserName) Then pUsername = !UserName
            If Not IsNull(!RankGrade) Then pRankGrade = !RankGrade
            If Not IsNull(!Role) Then pRole = !Role
            If Not IsNull(!MailAlert) Then pMailAlert = !MailAlert
            If Not IsNull(!Stations) Then pStations = !Stations
            If Not IsNull(!MessageRead) Then pMessageRead = !MessageRead
            
        End If
    End With
Set RstPerson = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstPerson As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstPerson = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE crewno = '" & pCrewNo & "' AND Deleted IS NULL")

    
    With RstPerson
        If .RecordCount = 0 Then
            .AddNew
        Else
            .Edit
        End If
        If pCrewNo <> "" Then !CrewNo = pCrewNo
        If pForename <> "" Then !Forename = pForename
        If pSurname <> "" Then !Surname = pSurname
        If pUsername <> "" Then !UserName = pUsername Else !UserName = pForename & " " & pSurname
        If pRankGrade <> "" Then !RankGrade = pRankGrade
        If pStations <> "" Then !Stations = pStations
        If pRole <> "" Then !Role = pRole
        !MailAlert = pMailAlert
        !MessageRead = pMessageRead
        .Update

    End With
    Set RstPerson = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstPerson As Recordset
    Dim i As Integer

    Set RstPerson = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE crewno = '" & pCrewNo & "' AND Deleted IS NULL")
    With RstPerson
        For i = .RecordCount To 1 Step -1
            If FullDelete Then
                .Delete
                .MoveNext
            Else
                .Edit
                !Deleted = Now
                .Update
            End If
        Next
    End With

    Set RstPerson = Nothing
End Sub

' ===============================================================
' Method LogOn
' Adds the user to the logged on table in DB
' ---------------------------------------------------------------
Public Sub LogOn()
    Dim RstUser As Recordset
    
    On Error GoTo ErrorHandler
    
    Set RstUser = SQLQuery("TblUsers")
    
    If RstUser.RecordCount = 0 Then
        DB.Execute "INSERT INTO TblUsers VALUES ('" & pUsername & "', '" & Now & "')"
    End If
Exit Sub

ErrorHandler:
    Err.Raise HANDLED_ERROR, , "Unable to add user to user list"
    Exit Sub
End Sub

' ===============================================================
' Method LogUserOff
' Removes the user from the logged on table in DB
' ---------------------------------------------------------------
Public Sub LogUserOff()

    On Error GoTo ErrorHandler
    
    DB.Execute "DELETE FROM TblUsers WHERE CurrentUsers = '" & pUsername & "'"
Exit Sub

ErrorHandler:
    Err.Raise HANDLED_ERROR, , "Unable to delete user from user list"
    Exit Sub
End Sub

' ===============================================================
Private Sub Class_Initialize()

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()

End Sub

'---------------------------------------------------------------


