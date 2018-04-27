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
    Initialise = True

Exit Function

ErrorExit:

    '***CleanUpCode***
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
