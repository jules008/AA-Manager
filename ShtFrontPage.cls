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
' Date - 27 Apr 18
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ShtFrontPage"

Private Agreement As ClsAgreement

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
    
    Set RstStation = Nothing
    
    Initialise = True

Exit Function

ErrorExit:

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
' BtnLoadAA
' Views AA for entered name
' ---------------------------------------------------------------
Private Sub BtnLoadAA()
    Dim SelectedPerson As TypeCrewMember
    
    Const StrPROCEDURE As String = "BtnLoadAA()"

    On Error GoTo ErrorHandler

    SelectedPerson = FrmPersonPicker.ShowForm
     
    If SelectedPerson.CrewNo <> "" Then
        With Agreement
            .CrewMember = SelectedPerson
            .DBGet
            .Display
        End With
    End If
Exit Sub

ErrorExit:

    '***CleanUpCode***
    Application.ScreenUpdating = True

Exit Sub

ErrorHandler:
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
    Agreement.Display
    Agreement.CreateNew
    
GracefulExit:

Exit Sub

ErrorExit:

    '***CleanUpCode***

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
    Application.ScreenUpdating = False

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
' Worksheet_Change(ByVal Target As Range)
' Detects a change to the template and formats accordingly
' ---------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim Cell As Range
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "Worksheet_Change(ByVal Target As Range)()"

    On Error GoTo ErrorHandler

Restart:
    If Agreement Is Nothing Then Err.Raise SYSTEM_RESTART
    
    Select Case Target.Cells.Count
        
        Case Is = 0
            Exit Sub
        
        Case Is = 1
            If Not Intersect(Target, Range("B4:AW98")) Is Nothing Then
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
    
    Set RngWk1 = Range(RNG_WEEK_1)

    For i = 0 To 11
        Me.Unprotect
        RngWk1.Offset(i * 8, 0).ClearContents
    Next

    TxtCrewNo = ""
    TxtName = ""
    TxtNoWeeks = ""
    TxtReviewDate = ""
    TxtTemplateDate = ""
    CmoContract = ""
    CmoRole = ""
    CmoStation1 = ""
    CmoStation2 = ""
    
    Me.Protect
    
    Clear = True
    
    Set RngWk1 = Nothing
Exit Function

ErrorExit:

    Me.Protect
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
    Dim i As Integer
    
    Const StrPROCEDURE As String = "FormatWkRnge()"

    On Error GoTo ErrorHandler
    
    Me.Unprotect
    
    Set RngWeek1 = ShtFrontPage.Range(RNG_WEEK_1)
        
    For i = 1 To NoWeeks
        Set RngWeek = RngWeek1.Offset((i * 8) - 8, 0)
        With RngWeek
            .Locked = False
            .Interior.Color = COLOUR_7
        End With
    Next
    
    For i = NoWeeks + 1 To 12
        Set RngWeek = RngWeek1.Offset((i * 8) - 8, 0)
        With RngWeek
            .Locked = True
            .Interior.Color = COLOUR_4
        End With
    Next

    TxtNoWeeks = NoWeeks
    
    FormatWkRnge = True
    
    Set RngWeek1 = Nothing
    Set RngWeek = Nothing
    Me.Protect
    
Exit Function

ErrorExit:

    Set RngWeek1 = Nothing
    Set RngWeek = Nothing
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

    Me.Protect
    
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
    
    If Intersect(RngSel, Range("B4:AW98")) Is Nothing Then Exit Sub
    
    If Selection.Cells.Count = 0 Then Exit Sub

    For Each Cell In Selection
        Cell.Value = Stn
    Next
    RngSel.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
GracefulExit:

Exit Sub

ErrorExit:

    Application.ScreenUpdating = False

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
