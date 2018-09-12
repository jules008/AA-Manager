VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAdminUserList 
   Caption         =   "Action Plan"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5280
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

Public Function ShowForm() As Boolean
    
   Const StrPROCEDURE As String = "ShowForm()"
   
   On Error GoTo ErrorHandler
   
    ResetForm
    If Not RefreshUserList Then Err.Raise HANDLED_ERROR
    Show

    ShowForm = True
Exit Function

ErrorExit:
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

Private Sub BtnClose_Click()
    On Error Resume Next
    
    Me.Hide
End Sub

Private Sub BtnCourseAdmin_Click()
    Const StrPROCEDURE As String = "BtnCourseAdmin_Click()"

    On Error GoTo ErrorHandler

    If Not FrmAdminUserList.ShowForm Then Err.Raise HANDLED_ERROR
Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
Private Sub BtnDelete_Click()
    
    Const StrPROCEDURE As String = "BtnDelete_Click()"
    
    Dim Response As Integer
    Dim SelUser As Integer
    Dim UserName As String
    
    On Error GoTo ErrorHandler
        
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
Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Private Sub BtnNew_Click()
    On Error Resume Next
    
    ResetForm
End Sub

Private Sub BtnUpdate_Click()

    Const StrPROCEDURE As String = "BtnUpdate_Click()"
       
    If Not AddUpdateUser(ActiveUser) Then Err.Raise HANDLED_ERROR
    
    If Not RefreshUserList Then Err.Raise HANDLED_ERROR
    
    If ValidateData = True Then
    End If
Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Private Sub CmoCourseNo_Change()
    Const StrPROCEDURE As String = "CmoCourseNo_Change()"

    On Error GoTo ErrorHandler

    If Not RefreshUserList Then Err.Raise HANDLED_ERROR
    
    ResetForm
Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
Private Sub LstAccessList_Click()
    Const StrPROCEDURE As String = "LstAccessList_Click()"

    On Error GoTo ErrorHandler

    If Not RefreshUserDetails Then Err.Raise HANDLED_ERROR
Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
Private Sub UserForm_Initialize()
    On Error Resume Next
    With LstHeadings
        .AddItem
        .List(0, 0) = "Users"
    End With
End Sub

Private Function ValidateData() As Boolean
    On Error Resume Next
    
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
End Function

Private Sub ResetForm()
    On Error Resume Next
    
    TxtCrewNo = ""
    TxtForeName = ""
    TxtRank = ""
    TxtSurname = ""
    ChkAdmin = False
End Sub

Public Function RefreshUserList() As Boolean
    Const StrPROCEDURE As String = "RefreshUserList()"

    Dim RstUserList As Recordset
    Dim i As Integer
    
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
    Set RstUserList = Nothing
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
Public Function RefreshUserDetails() As Boolean
    Const StrPROCEDURE As String = "RefreshUserDetails()"

    Dim ListSelection As Integer
    Dim UserName As String
    Dim RstUserDetails As Recordset
    
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
                TxtRank = !Rank
                TxtSurname = !Surname
                If !Admin = True Then ChkAdmin = True Else ChkAdmin = False
            End With
        End If
    End If
    Set RstUserDetails = Nothing
    RefreshUserDetails = True

Exit Function

ErrorExit:
    Set RstUserDetails = Nothing
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
