VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtFrontPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Class ShtFrontPage
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 04 Jun 18
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ShtFrontPage"

' ===============================================================
' Initialise
' Initialises sheet
' ---------------------------------------------------------------
Public Function Initialise() As Boolean
    Dim RstStation As Recordset
    
    Dim i As Integer
    
    Const StrPROCEDURE As String = "Initialise()"

    On Error GoTo ErrorHandler
    
    Set RstStation = ModDBLookups.ReturnStnList
    Set Agreement = Nothing
    Set Agreement = New ClsAgreement
    
    If Not FormatRevDate Then Err.Raise HANDLED_ERROR
    
    With RstStation
        CmoStation1.Clear
        CmoStation2.Clear
        Do While Not .EOF
            CmoStation1.AddItem !Callsign & " - " & !Name
            CmoStation2.AddItem !Callsign & " - " & !Name
            .MoveNext
            i = i + 1
        Loop
    End With
    
    With CmoNoWeeks
        .Clear
        .AddItem "Over 1 Week"
        For i = 2 To 12
            .AddItem "Over " & i & " Weeks"
        Next
    End With
    
    Me.Unprotect
    If Not DEV_MODE Then
        Me.Range("AZ:BA").EntireColumn.Hidden = True
    Else
        Me.Range("AZ:BA").EntireColumn.Hidden = False
    End If
    If PROTECT_ON Then Me.Protect
    
    With CmoContract
        .Clear
        .AddItem "Under 120Hrs"
        .AddItem "Over 120hrs"
    End With

    With CmoRole
        .Clear
        .AddItem "FF"
        .AddItem "CM"
        .AddItem "WM"
    End With

    If CurrentUser.Role = Admin Then
        Me.Shapes("BtnAdmin").Visible = msoTrue
    Else
        If DEV_MODE Then
            Me.Shapes("BtnAdmin").Visible = msoTrue
        Else
            Me.Shapes("BtnAdmin").Visible = msoFalse
        End If
    End If

    Set RstStation = Nothing
    
    Initialise = True

Exit Function

ErrorExit:

    If PROTECT_ON Then Me.Protect
    Set RstStation = Nothing
    Initialise = False

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
' FormatRevDate
' Formats the review date dependant on the number of days left
' ---------------------------------------------------------------
Public Function FormatRevDate() As Boolean
    Dim RevDaysLeft As Integer
    
    Const StrPROCEDURE As String = "FormatRevDate()"

    On Error GoTo ErrorHandler
    With TxtRevDate
        If IsDate(TxtRevDate) Then
            RevDaysLeft = DateDiff("d", Now, .Value)
            
            Select Case RevDaysLeft
                Case Is > 14
                    .BackColor = COLOUR_1
                    .ForeColor = COLOUR_5
                    
                Case Is > 0
                    .BackColor = COLOUR_9
                    .ForeColor = COLOUR_5
                Case Is < 0
                    .BackColor = COLOUR_10
                    .ForeColor = COLOUR_5
            End Select
        End If
    End With

    FormatRevDate = True


Exit Function

ErrorExit:

    '***CleanUpCode***
    FormatRevDate = False

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
' BtnLoadAA
' Views AA for entered name
' ---------------------------------------------------------------
Private Sub BtnLoadAA()
    Dim SelectedPerson As TypeCrewMember
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnLoadAA()"

    On Error GoTo ErrorHandler

Restart:

    SelectedPerson = FrmPersonPicker.ShowForm
    
    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If SelectedPerson.CrewNo <> "" Then
        
        Set Agreement = Nothing
        Set Agreement = New ClsAgreement
                
        With Agreement
            .CrewMember = SelectedPerson
            .DBGet
            .Clear
            .Display
        End With
'        TxtRevDate.Visible = True
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
' BtnNewAA
' Creates a new agreement
' ---------------------------------------------------------------
Private Sub BtnNewAA()
    Const StrPROCEDURE As String = "BtnNewAA()"

    On Error GoTo ErrorHandler

Restart:

    If Agreement Is Nothing Then Err.Raise SYSTEM_RESTART
    
    Set Agreement = Nothing
    Set Agreement = New ClsAgreement
    
    Application.ScreenUpdating = False
    
    Agreement.Clear
    Agreement.CreateNew
    
GracefulExit:
    
    Application.ScreenUpdating = True

Exit Sub

ErrorExit:

    Application.ScreenUpdating = True

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        If CustomErrorHandler(Err.Number) Then Resume Restart
    End If
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnUpdateAA
' Updates any changes to Agreement
' ---------------------------------------------------------------
Private Sub BtnUpdateAA()
    Dim ErrNo As Integer
    Dim Response As Integer
    
    Const StrPROCEDURE As String = "BtnUpdateAA()"

    On Error GoTo ErrorHandler
    
Restart:

    If Agreement Is Nothing Then Err.Raise SYSTEM_RESTART, , "Agreement Class is empty"

    Select Case ValidateForm
        
        Case Is = FunctionalError
            Err.Raise HANDLED_ERROR
            
        Case Is = ValidationError
            Exit Sub
                
    End Select
    
    Response = MsgBox("Are you sure you want to update the changes to the Agreement?  Once saved, the changes cannot be undone!", vbExclamation + vbYesNo, APP_NAME)
    
    If Response = 6 Then
        With Agreement
            .Update
            .DBSave
        End With
        MsgBox "The Agreement has been saved successfully", vbInformation, APP_NAME
    Else
        Response = MsgBox("Would you like to discard the changes and revert back to the saved Agreement?", vbYesNo + vbInformation, APP_NAME)
        
        If Response = 6 Then
            With Agreement
                .DBGet
                .Display
            End With
        End If
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

Private Sub TextBox5_Change()

End Sub

Private Sub TxtRevDate_Change()

End Sub

' ===============================================================
' Worksheet_Change(ByVal Target As Range)
' Detects a change to the template and formats accordingly
' ---------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim Cell As Range
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "Worksheet_Change(ByVal Target As Range)()"

    On Error GoTo ErrorHandler

Restart:
    
    Select Case Target.Cells.Count
        
        Case Is = 0
            Exit Sub
        
        Case Is = 1
            If Not Intersect(Target, Range("C4:AX98")) Is Nothing Then
                If Not FormatCells(Target) Then Err.Raise HANDLED_ERROR
            End If
        
        Case Is > 1
            Selection.Cells(1, 1).Select
            For Each Cell In Target
                If Not FormatCells(Cell) Then Err.Raise HANDLED_ERROR
            Next
    End Select
    
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
' Worksheet_Deactivate
' Clears Agreement class on exit
' ---------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    Set Agreement = Nothing
End Sub

' ===============================================================
' Clear
' Clears the AA form
' ---------------------------------------------------------------
Public Function Clear() As Boolean
    Dim RngWk1 As Range
    Dim i As Integer
    
    Const StrPROCEDURE As String = "Clear()"
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Set RngWk1 = Range(RNG_WEEK_1)
       
    For i = 0 To 11
        Me.Unprotect
        RngWk1.Offset(i * 9, 0).ClearContents
    Next
        
    If Not ShtFrontPage.FormatWkRnge(1) Then Err.Raise HANDLED_ERROR

    TxtCrewNo = ""
    TxtName = ""
    CmoNoWeeks = ""
    TxtRevDate.Visible = False
    TxtTemplateDate = ""
    CmoContract = ""
    CmoRole = ""
    CmoStation1 = ""
    CmoStation2 = ""
        
    If Not ShtFrontPage.UpdateValues Then Err.Raise HANDLED_ERROR

    If PROTECT_ON Then Me.Protect
    
    Clear = True

    Set RngWk1 = Nothing

    Application.ScreenUpdating = True

Exit Function

ErrorExit:
    
    Application.ScreenUpdating = True

    If PROTECT_ON Then Me.Protect
    Clear = False
    
    Set RngWk1 = Nothing
    
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
' FormatWkRnge
' sets and formats the correct number of weeks for the template
' ---------------------------------------------------------------
Public Function FormatWkRnge(NoWeeks As Integer) As Boolean
    Dim RngWeek1 As Range
    Dim RngWeek As Range
    Dim RngDay As Range
    Dim RngDay1 As Range
    Dim RngTime As Range
    Dim RngTime1 As Range
    Dim RngWkNo As Range
    Dim RngWkNo1 As Range
    
    Dim i As Integer
    
    Const StrPROCEDURE As String = "FormatWkRnge()"

    On Error GoTo ErrorHandler
    
    Me.Unprotect
    
    Set RngWeek1 = ShtFrontPage.Range(RNG_WEEK_1)
    Set RngTime1 = ShtFrontPage.Range(RNG_TIME_1)
    Set RngDay1 = ShtFrontPage.Range(RNG_DAY_1)
    Set RngWkNo1 = ShtFrontPage.Range(RNG_WKNO_1)
    
    For i = 1 To NoWeeks
    
        'Format timeslots
        Set RngWeek = RngWeek1.Offset((i * 9) - 9, 0)
        RngWeek.Select
        With RngWeek
            .Locked = False
            .Interior.Color = COLOUR_7
        End With
    
        'Format Times
        Set RngTime = RngTime1.Offset((i * 9) - 9, 0)
        RngTime.Select
        With RngTime
            .Locked = True
            .Interior.Color = COLOUR_8
            .Font.Color = COLOUR_2
        End With
    
        'Format Days
        Set RngDay = RngDay1.Offset((i * 9) - 9, 0)
        RngDay.Select
        With RngDay
            .Locked = True
            .Interior.Color = COLOUR_8
            .Font.Color = COLOUR_2
        End With
        
        'Format Week No
        Set RngWkNo = RngWkNo1.Offset((i * 9) - 9, 0)
        RngWkNo.Select
        With RngWkNo
            .Locked = True
            .Interior.Color = COLOUR_2
            .Font.Color = COLOUR_8
        End With
    Next
    
    For i = NoWeeks + 1 To 12
        Set RngWeek = RngWeek1.Offset((i * 9) - 9, 0)
        RngWeek.Select
        With RngWeek
            .Locked = True
            .Interior.Color = COLOUR_4
        End With
    
        'Format Times
        Set RngTime = RngTime1.Offset((i * 9) - 9, 0)
        RngTime.Select
        With RngTime
            .Locked = True
            .Interior.Color = COLOUR_4
            .Font.Color = COLOUR_4
        End With
    
        'Format Days
        Set RngDay = RngDay1.Offset((i * 9) - 9, 0)
        RngDay.Select
        With RngDay
            .Locked = True
            .Interior.Color = COLOUR_4
            .Font.Color = COLOUR_4
        End With
        
        'Format Week No
        Set RngWkNo = RngWkNo1.Offset((i * 9) - 9, 0)
        RngWkNo.Select
        With RngWkNo
            .Locked = True
            .Interior.Color = COLOUR_4
            .Font.Color = COLOUR_4
        End With
    Next

    Me.Range("C4").Select
            
    FormatWkRnge = True
    
    Set RngWeek1 = Nothing
    Set RngWeek = Nothing
    Set RngDay = Nothing
    Set RngTime = Nothing
    Set RngDay1 = Nothing
    Set RngTime1 = Nothing
    Set RngWkNo = Nothing
    Set RngWkNo1 = Nothing
    
    If PROTECT_ON Then Me.Protect
    
Exit Function

ErrorExit:

    Set RngWeek1 = Nothing
    Set RngWeek = Nothing
    Set RngDay = Nothing
    Set RngTime = Nothing
    Set RngDay1 = Nothing
    Set RngTime1 = Nothing
    Set RngWkNo = Nothing
    Set RngWkNo1 = Nothing
        
    Me.Protect
    
    FormatWkRnge = False

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
' FormatCells
' Formats cells containing 1 and 2 to the correct colours
' ---------------------------------------------------------------
Private Function FormatCells(RngTarget As Range) As Boolean
    Const StrPROCEDURE As String = "FormatCells()"

    On Error GoTo ErrorHandler

    Me.Unprotect
    
    Select Case RngTarget
        Case 0
             With RngTarget
                .Interior.Color = COLOUR_7
                .Font.Color = COLOUR_7
            End With
            
           
        Case 1
            With RngTarget
                .Interior.Color = COLOUR_1
                .Font.Color = COLOUR_2
            End With
            
        Case 2
            With RngTarget
                .Interior.Color = COLOUR_2
                .Font.Color = COLOUR_1
            End With
    End Select
    FormatCells = True

    If PROTECT_ON Then Me.Protect
    
Exit Function

ErrorExit:
    Me.Protect
    
    FormatCells = False

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
' BtnFillStn0
' Fills selected cells with 0
' ---------------------------------------------------------------
Private Sub BtnFillStn0()
    FillStnCells (0)
End Sub

' ===============================================================
' BtnAdmin
' Shows Admin Screen
' ---------------------------------------------------------------
Private Sub BtnAdmin()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "ShowAdmin()"

    On Error GoTo ErrorHandler

Restart:

    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART

    If Not FrmAdminUserList.ShowForm Then Err.Raise HANDLED_ERROR
    
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
' BtnFillStn1
' Fills selected cells with 1
' ---------------------------------------------------------------
Private Sub BtnFillStn1()
    FillStnCells (1)
End Sub

' ===============================================================
' BtnFillStn2
' Fills selected cells with 2
' ---------------------------------------------------------------
Private Sub BtnFillStn2()
    FillStnCells (2)
End Sub

' ===============================================================
' FillStnCells
' Fills selected cells with 0, 1 or 2
' ---------------------------------------------------------------
Private Sub FillStnCells(Stn As Integer)
    Dim ErrNo As Integer
    Dim Cell As Range
    Dim RngSel As Range
    
    Const StrPROCEDURE As String = "FillStnCells()"

    On Error GoTo ErrorHandler

Restart:
    
    Application.ScreenUpdating = False
    
    Set RngSel = Selection
    Debug.Print RngSel.Address
    
    If Agreement Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If Intersect(RngSel, Range("C4:AX98")) Is Nothing Then Exit Sub
    
    If Selection.Cells.Count = 0 Then Exit Sub

    For Each Cell In Selection
        Cell.Value = Stn
    Next
    RngSel.Cells(1, 1).Select
    
    If Not UpdateValues Then Err.Raise HANDLED_ERROR
    
GracefulExit:

    Application.ScreenUpdating = True

Exit Sub

ErrorExit:

    Application.ScreenUpdating = True
    Set RngSel = Nothing
    
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
' UpdateValues
' Updates the hrs/wk totals
' ---------------------------------------------------------------
Public Function UpdateValues() As Boolean
    Const StrPROCEDURE As String = "UpdateValues()"

    On Error GoTo ErrorHandler
    Me.Unprotect
    
    If CmoNoWeeks.ListIndex <> -1 Then
    
        [NoWeeks] = CmoNoWeeks.ListIndex + 1
        
        If PROTECT_ON Then Me.Protect
        
        Me.TxtStn1Hrs.Value = Format([Stn1Total] / 2 / [NoWeeks], "0.0")
        Me.TxtStn2Hrs.Value = Format([Stn2Total] / 2 / [NoWeeks], "0.0")
        Me.TxtTotalHrs.Value = Format([AllStnTotal] / 2 / [NoWeeks], "0.0")
    
        With CmoContract
            If TxtTotalHrs < 120 Then .ListIndex = 0 Else .ListIndex = 1
        End With
    Else
        Me.TxtStn1Hrs.Value = Format(0, "0")
        Me.TxtStn2Hrs.Value = Format(0, "0")
        Me.TxtTotalHrs.Value = Format(0, "0")
    End If
    
    Me.TxtRevDate = Format(Agreement.RevDateDue, "dd mmm yy")
            
    If Not Me.FormatRevDate Then Err.Raise HANDLED_ERROR
    
    UpdateValues = True


Exit Function

ErrorExit:

    '***CleanUpCode***
    UpdateValues = False

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
' ValidateForm
' Validates input form before updating database
' ---------------------------------------------------------------
Private Function ValidateForm() As EnumFormValidation
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler

    If CmoContract = -1 Then
        ValidateForm = ValidationError
        MsgBox "Please select a Contract", vbExclamation + vbOKOnly, APP_NAME
        Exit Function
    End If
    
    If CmoNoWeeks.ListIndex = -1 Then
        ValidateForm = ValidationError
        MsgBox "Please select the number of weeks for the Contract", vbExclamation + vbOKOnly, APP_NAME
        Exit Function
    End If

    If CmoRole.ListIndex = -1 Then
        ValidateForm = ValidationError
        MsgBox "Please select a Role", vbExclamation + vbOKOnly, APP_NAME
        Exit Function
    End If

    If CmoStation1.ListIndex = -1 Then
        ValidateForm = ValidationError
        MsgBox "Please select a primary Station", vbExclamation + vbOKOnly, APP_NAME
        Exit Function
    End If

    If CmoStation2.ListIndex = -1 And TxtStn2Hrs <> 0 Then
        ValidateForm = ValidationError
        MsgBox "Please select a second Station", vbExclamation + vbOKOnly, APP_NAME
        Exit Function
    End If

    If TxtCrewNo = -1 Then
        ValidateForm = ValidationError
        MsgBox "Please enter a Crew No", vbExclamation + vbOKOnly, APP_NAME
        Exit Function
    End If

    If Not IsNumeric(TxtCrewNo) Then
        ValidateForm = ValidationError
        MsgBox "Please only use numerics for the Crew No", vbExclamation + vbOKOnly, APP_NAME
        Exit Function
    End If
    
    If Len(TxtCrewNo) > 4 Then
        ValidateForm = ValidationError
        MsgBox "Please enter a valid Crew No", vbExclamation + vbOKOnly, APP_NAME
        Exit Function
    End If
    
    If TxtName = "" Then
        ValidateForm = ValidationError
        MsgBox "Please enter a name", vbExclamation + vbOKOnly, APP_NAME
        Exit Function
    End If
    
    ValidateForm = FormOK

Exit Function

ErrorExit:

    ValidateForm = FunctionalError

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
' BtnApprove
' Approves Agreement and extends date by 6 months
' ---------------------------------------------------------------
Private Sub BtnApprove()
    Dim ErrNo As Integer
    Dim Response As Integer
    
    Const StrPROCEDURE As String = "BtnApprove()"

    On Error GoTo ErrorHandler

Restart:

    If Agreement Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If Agreement.CrewMember.Name <> "" Then
        Response = MsgBox("I have reviewed this Availability Agreement with " & Agreement.CrewMember.Name _
                 & " and we agree that this agreement is achievable and realistic over the next six months", vbYesNo + vbExclamation, APP_NAME)
                
        If Response = 6 Then
            Agreement.RevDateDue = DateAdd("d", 180, Now)
            Agreement.DBSave
            
            TxtRevDate.Visible = True
            
            If Not Me.UpdateValues Then Err.Raise HANDLED_ERROR
            
        End If
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
