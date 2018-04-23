VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsAgreement
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Apr 18
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
'===============================================================

Option Explicit
Private pCrewNo As String
Private pTimeSlot(1 To 84, 1 To 48) As Integer
Private pCrewMember As TypeCrewMember
Private pAADate As Date
Private pContractType As Integer
Private pHrsPW As Integer
Private pNoWeeks As Integer
Private pStation1 As TypeStation
Private pStation2 As TypeStation
Private pRevDateDue As Date

'---------------------------------------------------------------
Public Property Get CrewNo() As String
    CrewNo = pCrewMember.CrewNo
End Property

Public Property Let CrewNo(ByVal vNewValue As String)
    pCrewMember.CrewNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TimeSlot(Day As Integer, Slot As Integer) As Integer
    TimeSlot = pTimeSlot(Day, Slot)
End Property

Public Property Let TimeSlot(Day As Integer, Slot As Integer, ByVal vNewValue As Integer)
    pTimeSlot(Day, Slot) = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CrewMember() As TypeCrewMember
    CrewMember = pCrewMember
End Property

Public Property Let CrewMember(ByRef vNewValue As TypeCrewMember)
    pCrewMember = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AADate() As Date
    AADate = pAADate
End Property

Public Property Let AADate(ByVal vNewValue As Date)
    pAADate = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ContractType() As Integer
    ContractType = pContractType
End Property

Public Property Let ContractType(ByVal vNewValue As Integer)
    pContractType = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get HrsPW() As Integer
    HrsPW = pHrsPW
End Property

Public Property Let HrsPW(ByVal vNewValue As Integer)
    pHrsPW = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get NoWeeks() As Integer
    NoWeeks = pNoWeeks
End Property

Public Property Let NoWeeks(ByVal vNewValue As Integer)
    pNoWeeks = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Station1() As TypeStation
    Station1 = pCrewMember.Station1
End Property

Public Property Let Station1(ByRef vNewValue As TypeStation)
    pCrewMember.Station1 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Station2() As TypeStation
    Station2 = pCrewMember.Station2
End Property

Public Property Let Station2(ByRef vNewValue As TypeStation)
    pCrewMember.Station2 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get RevDateDue() As Date
    RevDateDue = pRevDateDue
End Property

Public Property Let RevDateDue(ByVal vNewValue As Date)
    pRevDateDue = vNewValue
End Property

'---------------------------------------------------------------
' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet()
    Dim RstAgreement As Recordset
    Dim RstSlots As Recordset

    Set RstAgreement = ModDatabase.SQLQuery("SELECT * FROM TblAgreement WHERE CrewNo = " & pCrewNo)
    Set RstSlots = ModDatabase.SQLQuery("SELECT * FROM TemplateDetail WHERE CrewNo = " & pCrewNo)
    
    With RstAgreement
        If .RecordCount > 0 Then
            If Not IsNull(!CrewNo) Then pCrewMember.CrewNo = !CrewNo
            If Not IsNull(!CrewName) Then pCrewMember.Name = !CrewName
            If Not IsNull(!TemplateDate) Then pAADate = !TemplateDate
            If Not IsNull(!ContractType) Then pContractType = !ContractType
            If Not IsNull(!HrsPW) Then pHrsPW = !HrsPW
            If Not IsNull(!NoWeeks) Then pNoWeeks = !NoWeeks
            If Not IsNull(!RevDateDue) Then pRevDateDue = !RevDateDue
        End If
        
        If .RecordCount = 1 Then
            If Not IsNull(!Station1) Then pStation1.StationNo = !Station1
        End If
        
        If .RecordCount = 2 Then
            Do While Not .EOF
            
                If !NoStation = 1 Then
                    If Not IsNull(!Station1) Then pStation1.StationNo = !Station1
                Else
                    If Not IsNull(!Station2) Then pStation2.StationNo = !Station2
                End If
            Loop
        End If
    End With
            
    With RstSlots
        If .RecordCount > 0 Then
            Do While Not .EOF
                pTimeSlot(!Day, !TimeSlot) = !oncall
                .MoveNext
            Loop
        End If
    End With
    Set RstAgreement = Nothing
    Set RstSlots = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstAgreement As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstAgreement = ModDatabase.SQLQuery("SELECT * FROM TblAgreement WHERE CrewNo = " & pCrewNo & " AND Deleted IS NULL")
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(CrewNo) FROM TblAgreement ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstAgreement
        If .RecordCount = 0 Then
            .addnew
            pCrewNo = LastNo + 1
        Else
            .Edit
        End If
'        !CrewNo = pCrewNo
'        !TimeSlot = pTimeSlot
'        !CrewMember = pCrewMember
'        !AADate = pAADate
'        !ContractType = pContractType
'        !HrsPW = pHrsPW
'        !NoWeeks = pNoWeeks
'        !Station = pStation
'        !RevDateDue = pRevDateDue
'        .Update

    End With
    Set RstAgreement = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstAgreement As Recordset
    Dim i As Integer

    Set RstAgreement = ModDatabase.SQLQuery("SELECT * FROM TblAgreement WHERE CrewNo = " & pCrewNo & " AND Deleted IS NULL")
    With RstAgreement
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

    Set RstAgreement = Nothing
End Sub


' ===============================================================
Private Sub Class_Initialize()

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()

End Sub

'---------------------------------------------------------------
