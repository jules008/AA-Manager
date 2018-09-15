VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAdminUserList 
   Caption         =   "Action Plan"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12150
   OleObjectBlob   =   "FrmAdminUserList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAdminUserList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 09 Nov 16
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmAdminUserList"

' ===============================================================
' ShowForm
' Displays form
' ---------------------------------------------------------------
Public Function ShowForm() As Boolean
    Const StrPROCEDURE As String = "ShowForm()"

    On Error GoTo ErrorHandler

    If Not ResetForm Then Err.Raise HANDLED_ERROR
    
    If Not RefreshUserList Then Err.Raise HANDLED_ERROR
    Show

    ShowForm = True
Exit Function

ErrorExit:

    '***CleanUpCode***
    ShowForm = False

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
' BtnClose_Click
' Closes form
' ---------------------------------------------------------------
Private Sub BtnClose_Click()
    On Error Resume Next
    
    Me.Hide
End Sub

' ===============================================================
' BtnDelete_Click
' Deletes person from list
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim ErrNo As Integer
    Dim Response As Integer
    Dim SelUser As Integer
    Dim UserName As String

    Const StrPROCEDURE As String = "BtnDelete_Click()"
    On Error GoTo ErrorHandler

Restart:
    
    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART
    
    SelUser = LstAccessList.ListIndex
    
    If SelUser <> -1 Then
        UserName = LstAccessList.List(SelUser, 0)
        Response = MsgBox("Are you sure you want to remove " _
                            & UserName & " from the system? ", 36)
    
        If Response = 6 Then
            If Not ModSecurity.RemoveUser(UserName) Then Err.Raise HANDLED_ERROR
        End If
        If Not RefreshUserList Then Err.Raise HANDLED_ERROR
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
    End If


GracefulExit:


Exit Sub

ErrorExit:

    '***CleanUpCode***

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnNew_Click
' Creates new person in list
' ---------------------------------------------------------------
Private Sub BtnNew_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnNew_Click()"

    On Error GoTo ErrorHandler

Restart:

    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If Not ResetForm Then Err.Raise HANDLED_ERROR

GracefulExit:


Exit Sub

ErrorExit:

    '***CleanUpCode***

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnUpdate_Click
' Updates changes to database
' ---------------------------------------------------------------
Private Sub BtnUpdate_Click()
    Dim User As ClsPerson
    Dim ErrNo As Integer
    Dim StrStations As String
    Dim i As Integer
    Dim Response As Integer
    Dim Cntrl As Control
    Dim Stations(1 To 38) As String

    Const StrPROCEDURE As String = "BtnUpdate_Click()"

    On Error GoTo ErrorHandler

Restart:
    
    Set User = New ClsPerson
    
    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART

    If ValidateData = True Then
    
        For i = 1 To 38
            Set Cntrl = Me.Controls("ChkStn" & i)
            If Cntrl Then Stations(i) = 1 Else Stations(i) = 0
        Next
        
        With User
            .CrewNo = TxtCrewNo
            .Forename = TxtForeName
            .RankGrade = TxtRank
            .Role = CmoRole.ListIndex
            .Stations = Join(Stations, ";")
            .Surname = TxtSurname
            .UserName = TxtUserName
        End With
        
        If Not AddUpdateUser(User) Then Err.Raise HANDLED_ERROR
        
        If Not RefreshUserList Then Err.Raise HANDLED_ERROR
        
        MsgBox "The record has been updated", vbOKOnly + vbInformation, APP_NAME

    End If

GracefulExit:

    Set User = Nothing
    Set Cntrl = Nothing

Exit Sub

ErrorExit:

    Set User = Nothing
    Set Cntrl = Nothing
    '***CleanUpCode***

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub



' ===============================================================
' LstAccessList_Click
' Event triggered when name in list is selected
' ---------------------------------------------------------------
Private Sub LstAccessList_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "LstAccessList_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Not PopulateForm Then Err.Raise HANDLED_ERROR

GracefulExit:


Exit Sub

ErrorExit:

    '***CleanUpCode***

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' FormInitialise
' Initialisation routine when form starts up
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Dim Roles(1 To 3, 1 To 2) As String
    
    Const StrPROCEDURE As String = "FormInitialise()"

    On Error GoTo ErrorHandler

    With LstHeadings
        .Clear
        .AddItem
        .List(0, 0) = "Users"
    End With
    
    Roles(1, 1) = 0
    Roles(1, 2) = "WCS"
    Roles(2, 1) = 1
    Roles(2, 2) = "FDS"
    Roles(3, 1) = 2
    Roles(3, 2) = "Admin"
    
    With CmoRole
        .Clear
        .List() = Roles
    End With
    
    FormInitialise = True


Exit Function

ErrorExit:

    '***CleanUpCode***
    FormInitialise = False

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
' UserForm_Initialize
' Initialisation routine when form starts up
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "UserForm_Initialize()"

    On Error GoTo ErrorHandler

Restart:

    If Not FormInitialise Then Err.Raise HANDLED_ERROR

GracefulExit:

Exit Sub

ErrorExit:

    '***CleanUpCode***

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' ValidateData
' Validates input user data
' ---------------------------------------------------------------
Private Function ValidateData() As Boolean
    Const StrPROCEDURE As String = "ValidateData()"

    On Error GoTo ErrorHandler
    
    If Me.TxtForeName = "" Then
        MsgBox "Please enter the User's forename"
        ValidateData = False
        Exit Function
    End If
    
    If Me.TxtSurname = "" Then
        MsgBox "Please enter the User's surname"
        ValidateData = False
        Exit Function
    End If
    
    If Me.CmoRole.ListIndex = -1 Then
        MsgBox "Please select the User's role"
        ValidateData = False
        Exit Function
    End If

    ValidateData = True


Exit Function

ErrorExit:

    '***CleanUpCode***
    ValidateData = False

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
' ResetForm
' Resets all controls on form
' ---------------------------------------------------------------
Private Function ResetForm() As Boolean
    Dim i As Integer
    Dim ChkBox As MSForms.CheckBox
    
    Const StrPROCEDURE As String = "ResetForm()"

    On Error GoTo ErrorHandler

    TxtCrewNo = ""
    TxtForeName = ""
    TxtRank = ""
    TxtSurname = ""
    CmoRole = ""
    
    For i = 1 To 38
        Set ChkBox = Me.Controls("ChkStn" & i)
        ChkBox = False
    Next

    ResetForm = True


Exit Function

ErrorExit:

    '***CleanUpCode***
    ResetForm = False

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
' RefreshUserList
' Refreshes list of users
' ---------------------------------------------------------------
Private Function RefreshUserList() As Boolean
    Dim RstUserList As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "RefreshUserList()"
    
    On Error GoTo ErrorHandler

   Set RstUserList = GetAccessList
    
    LstAccessList.Clear
    
    If Not RstUserList Is Nothing Then
        With RstUserList
            Do While Not .EOF
                    
                LstAccessList.AddItem
                LstAccessList.List(i, 0) = RstUserList!UserName
                .MoveNext
                i = i + 1
            Loop
        End With
    End If
    Set RstUserList = Nothing

    RefreshUserList = True
Exit Function

ErrorExit:

    '***CleanUpCode***
    RefreshUserList = False

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
' PopulateForm
' Refreshes user details on form
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim ListSelection As Integer
    Dim UserName As String
    Dim Stations() As String
    Dim RstUserDetails As Recordset
    Dim Ctrl As Control
    Dim i As Integer
    
    Const StrPROCEDURE As String = "PopulateForm()"
    
    On Error GoTo ErrorHandler

    ListSelection = LstAccessList.ListIndex
    
    If ListSelection = -1 Then
        If Not ResetForm Then Err.Raise HANDLED_ERROR
    Else
        UserName = LstAccessList.List(ListSelection, 0)
        Set RstUserDetails = GetUserDetails(UserName)
        
        If Not RstUserDetails Is Nothing Then
            With RstUserDetails
                TxtCrewNo = !CrewNo
                TxtForeName = !Forename
                TxtRank = !RankGrade
                TxtSurname = !Surname
                CmoRole.ListIndex = !Role
                TxtUserName = !UserName
                Stations = Split(!Stations, ";")
            End With
            
            For i = 0 To 37
                Set Ctrl = Me.Controls("ChkStn" & i + 1)
                If Stations(i) = 1 Then Ctrl = True Else Ctrl = False
            Next
        End If
    End If
    
    Set Ctrl = Nothing
    Set RstUserDetails = Nothing
    PopulateForm = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    Set Ctrl = Nothing
    Set RstUserDetails = Nothing
    PopulateForm = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
