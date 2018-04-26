Attribute VB_Name = "ModDBLookups"
'===============================================================
' Module ModDBLookups
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 26 Apr 18
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModDBLookups"


' ===============================================================
' StationLookUp
' Returns Station type from Callsign, Name or number
' ---------------------------------------------------------------
Public Function StationLookUp(Optional Callsign As String, Optional StationNo As Integer, Optional StationName As String) As TypeStation
    Dim RstStation As Recordset
    Dim StationRecord As TypeStation
    Dim SQLStr As String
    
    Const StrPROCEDURE As String = "StationLookUp()"

    On Error GoTo ErrorHandler

    If Callsign <> "" Then
        SQLStr = "SELECT * FROM TblStnLookUp WHERE CallSign = '" & Callsign & "'"
    Else
        If StationNo <> 0 Then
            SQLStr = "SELECT * FROM TblStnLookUp WHERE StationNo = " & StationNo
        Else
            If StationName <> "" Then
                SQLStr = "SELECT * FROM TblStnLookUp WHERE Name = '" & StationName & "'"
            End If
        End If
    End If
    
    If SQLStr <> "" Then
        Set RstStation = SQLQuery(SQLStr)
        
        With RstStation
            If .RecordCount > 0 Then
                With StationRecord
                    .StationCallSign = RstStation!Callsign
                    .StationName = RstStation!Name
                    .StationNo = RstStation!StationNo
                End With
            End If
        End With
    End If
    
    StationLookUp = StationRecord
    Set RstStation = Nothing
Exit Function

ErrorExit:

'    ***CleanUpCode***
    StationRecord.StationNo = 99
    StationLookUp = StationRecord
    Set RstStation = Nothing

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
' ContractLookUp
' Returns Contract Type from Number
' ---------------------------------------------------------------
Public Function ContractLookUp(ContractNo) As String
    Dim RstContract As Recordset
    
    Const StrPROCEDURE As String = "ContractLookUp()"

    On Error GoTo ErrorHandler

    Set RstContract = SQLQuery("SELECT * FROM TblContractLookUp WHERE ContractNo = " & ContractNo)
        
    With RstContract
        If .RecordCount > 0 Then
            ContractLookUp = !ContractType
        End If
    End With
    
    Set RstContract = Nothing
Exit Function

ErrorExit:

    Set RstContract = Nothing

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
' ReturnStnList
' Returns a list of stations in a recordset
' ---------------------------------------------------------------
Public Function ReturnStnList() As Recordset
    Dim RstStations As Recordset
    
    Const StrPROCEDURE As String = "ReturnStnList()"

    On Error GoTo ErrorHandler

    Set RstStations = SQLQuery("SELECT StationNo, Callsign, Name FROM TblStnLookUp")
    
    Set ReturnStnList = RstStations

    Set RstStations = Nothing
Exit Function

ErrorExit:

    Set RstStations = Nothing
    Set ReturnStnList = Nothing

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
