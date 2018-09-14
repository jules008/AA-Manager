VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAdminUserList 
   Caption         =   "Action Plan"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12540
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

Private ActiveUser As ClsPerson

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
    
    SelUser = LstAccessList.ListIndex
    
    If SelUser <> -1 Then
        UserName = LstAccessList.List(SelUser, 0)
        Response = MsgBox("Are you sure you want to remove " _
                            & UserName & " from the system? ", 36)
    
        If Response = 6 Then
            If Not ModSecurity.RemoveUser(UserName) Then Err.Raise HANDLED_ERROR
        End If
        If Not RefreshUserList Then Err.Raise HANDLED_ERROR
        If Not RefreshUserDetails Then Err.Raise HANDLED_ERROR
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
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnUpdate_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Not AddUpdateUser(ActiveUser) Then Err.Raise HANDLED_ERROR
    
    If Not RefreshUserList Then Err.Raise HANDLED_ERROR
    
    If ValidateData = True Then
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

Private Sub Label36_Click()

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

    If Not RefreshUserDetails Then Err.Raise HANDLED_ERROR

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
    Const StrPROCEDURE As String = "FormInitialise()"

    On Error GoTo ErrorHandler

    With LstHeadings
        .Clear
        .AddItem
        .List(0, 0) = "Users"
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
    
    If Me.TxtCrewNo = "" Then
        MsgBox "Please enter the User's Crew No"
        ValidateData = False
        Exit Function
    End If
    
    If Me.TxtForeName = "" Then
        MsgBox "Please enter the User's forename"
        ValidateData = False
        Exit Function
    End If
    
    If Me.TxtRank = "" Then
        MsgBox "Please enter the User's Rank"
        ValidateData = False
        Exit Function
    End If
    
    If Me.TxtSurname = "" Then
        MsgBox "Please enter the User's surname"
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
    ChkAdmin = False
    
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
' RefreshUserDetails
' Refreshes user details on form
' ---------------------------------------------------------------
Private Function RefreshUserDetails() As Boolean
    Dim ListSelection As Integer
    Dim UserName As String
    Dim RstUserDetails As Recordset
    
    Const StrPROCEDURE As String = "RefreshUserDetails()"
    
    On Error GoTo ErrorHandler

    ListSelection = LstAccessList.ListIndex
    
    If ListSelection = -1 Then
        TxtCrewNo = ""
        TxtForeName = ""
        TxtRank = ""
        TxtSurname = ""
        ChkAdmin = False
    Else
        UserName = LstAccessList.List(ListSelection, 0)
        Set RstUserDetails = GetUserDetails(UserName)
        
        If Not RstUserDetails Is Nothing Then
            With RstUserDetails
                TxtCrewNo = !CrewNo
                TxtForeName = !Forename
                TxtRank = !RankGrade
                TxtSurname = !Surname
                If !Admin = True Then ChkAdmin = True Else ChkAdmin = False
            End With
        End If
    End If
    Set RstUserDetails = Nothing
    RefreshUserDetails = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    RefreshUserDetails = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
