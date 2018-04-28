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
            CmoStation1.AddItem
            CmoStation2.AddItem
            CmoStation1.List(i, 0) = CStr(!StationNo)
            CmoStation1.List(i, 1) = !Callsign & " - " & !Name
            CmoStation2.List(i, 0) = CStr(!StationNo)
            CmoStation2.List(i, 1) = !Callsign & " - " & !Name
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
' BtnLoadAA_Click
' Views AA for entered name
' ---------------------------------------------------------------
Private Sub BtnLoadAA_Click()
    Dim SelectedPerson As TypeCrewMember
    
    Const StrPROCEDURE As String = "BtnLoadAA_Click()"

    On Error GoTo ErrorHandler

    SelectedPerson = FrmPersonPicker.ShowForm
     
    If SelectedPerson.CrewNo = "" Then err.Raise HANDLED_ERROR, , "No person returned from Person Picker Form"
    
    With Agreement
        .CrewMember = SelectedPerson
        .DBGet
        .Display
    End With
    
Exit Sub

ErrorExit:

    '***CleanUpCode***

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
' BtnNewAA_Click
' Creates a new agreement
' ---------------------------------------------------------------
Private Sub BtnNewAA_Click()
    Const StrPROCEDURE As String = "BtnNewAA_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Agreement Is Nothing Then err.Raise SYSTEM_RESTART
    
    Agreement.CreateNew
    
GracefulExit:

Exit Sub

ErrorExit:

    '***CleanUpCode***

Exit Sub

ErrorHandler:
    If err.Number >= 1000 And err.Number <= 1500 Then
        If CustomErrorHandler(err.Number) Then Resume Restart
    End If
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnUpdateAA_Click
' Updates any changes to Agreement
' ---------------------------------------------------------------
Private Sub BtnUpdateAA_Click()
    Const StrPROCEDURE As String = "BtnUpdateAA_Click()"

    On Error GoTo ErrorHandler

    If Agreement Is Nothing Then err.Raise GENERIC_ERROR, , "Agreement Class is empty"

    With Agreement
        .Update
        .DBSave
    End With
    
    MsgBox ""
Exit Sub

ErrorExit:

    '***CleanUpCode***

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
    
    Clear = True
    
    Set RngWk1 = Nothing
Exit Function

ErrorExit:

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
' Formats the given week range correctly and sets protection
' ---------------------------------------------------------------
Public Function FormatWkRnge(NoWeeks As Integer) As Boolean
    Dim RngWeek1 As Range
    Dim RngWeek As Range
    Dim i As Integer
    
    Const StrPROCEDURE As String = "FormatWkRnge()"

    On Error GoTo ErrorHandler
    
    Set RngWeek1 = ShtFrontPage.Range(RNG_WEEK_1)
        
    For i = 1 To NoWeeks
        Set RngWeek = RngWeek1.Offset((i * 8) - 8, 0)
        With RngWeek
            .Locked = True
            .Interior.Color = COLOUR_1
        End With
    Next
    
    For i = NoWeeks + 1 To 12
        Set RngWeek = RngWeek1.Offset((i * 8) - 8, 0)
        With RngWeek
            .Locked = False
            .Interior.Color = COLOUR_2
        End With
    Next

    FormatWkRnge = True
    
    Set RngWeek1 = Nothing
    Set RngWeek = Nothing

Exit Function

ErrorExit:

    Set RngWeek1 = Nothing
    Set RngWeek = Nothing
    
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
