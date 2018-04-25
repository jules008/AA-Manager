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
' Date - 25 Apr 18
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
'===============================================================

Option Explicit
Private pCrewNo As String
Private pTimeSlot(1 To 48, 1 To 84) As Integer
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
    TimeSlot = pTimeSlot(Slot, Day)
End Property

Public Property Let TimeSlot(Day As Integer, Slot As Integer, ByVal vNewValue As Integer)
    pTimeSlot(Slot, Day) = vNewValue
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
    
    Set RstAgreement = ModDatabase.SQLQuery("SELECT * FROM TblTemplate WHERE CrewNo = '" & pCrewMember.CrewNo & "'")
    Set RstSlots = ModDatabase.SQLQuery("SELECT * FROM TblTemplateDetail WHERE CrewNo = '" & pCrewMember.CrewNo & "'")
    
    With RstAgreement
        If .RecordCount > 0 Then
            If Not IsNull(!CrewNo) Then pCrewMember.CrewNo = !CrewNo
            If Not IsNull(!CrewName) Then pCrewMember.Name = !CrewName
            If Not IsNull(!Role) Then pCrewMember.Role = !Role
            If Not IsNull(!TemplateDate) Then pAADate = !TemplateDate
            If Not IsNull(!ContractType) Then pContractType = !ContractType
            If Not IsNull(!HrsPW) Then pHrsPW = !HrsPW
            If Not IsNull(!NoWeeks) Then pNoWeeks = !NoWeeks
            If Not IsNull(!RevDateDue) Then pRevDateDue = !RevDateDue
        End If
        
        If .RecordCount = 1 Then
            If Not IsNull(!StationNo) Then pStation1 = ModDBLookups.StationLookUp(StationNo:=!StationNo)
        End If
        
        If .RecordCount = 2 Then
            Do While Not .EOF
            
                If !NoStation = 1 Then
                    If Not IsNull(!StationNo) Then pStation1 = ModDBLookups.StationLookUp(StationNo:=!StationNo)
                Else
                    If Not IsNull(!StationNo) Then pStation2 = ModDBLookups.StationLookUp(StationNo:=!StationNo)
                End If
            Loop
        End If
    End With
            
    With RstSlots
        If .RecordCount > 0 Then
            Do While Not .EOF
                pTimeSlot(!TimeSlot, !Day) = !OnCall
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
    Dim RstSlots As Recordset
    Dim i As Integer
    Dim x As Integer
    Dim w As Integer

    Set RstAgreement = ModDatabase.SQLQuery("SELECT * FROM TblTemplate WHERE CrewNo = '" & pCrewMember.CrewNo & "'")
    Set RstSlots = ModDatabase.SQLQuery("TblTemplateDetail")
    
    With RstAgreement
        If .RecordCount > 0 Then
            .Edit
            !CrewNo = pCrewMember.CrewNo
            !CrewName = pCrewMember.Name
            !Role = pCrewMember.Role
            !TemplateDate = pAADate
            !ContractType = pContractType
            !HrsPW = pHrsPW
            !NoWeeks = pNoWeeks
            !RevDateDue = pRevDateDue
            .Update
        End If
        
        If .RecordCount = 1 Then
            .Edit
            !StationNo = pStation1.StationNo
            .Update
        End If
        
        If .RecordCount > 1 Then
            Do While Not .EOF
                If !NoStation = 1 Then
                    .Edit
                    !StationNo = pStation1.StationNo
                    .Update
                Else
                    .Edit
                    !StationNo = pStation2.StationNo
                    .Update
                End If
                .MoveNext
            Loop
        End If
    End With
    
    DB.Execute ("DELETE FROM TblTemplateDetail WHERE CrewNo = '" & pCrewMember.CrewNo & "'")
    
    With RstSlots
        For w = 1 To pNoWeeks
            For i = 1 To 7
                For x = 1 To 48
                    .AddNew
                    !CrewNo = pCrewMember.CrewNo
                    !Day = i + (w * 7) - 7
                    !Weekday = i
                    !TimeSlot = x
                    !OnCall = pTimeSlot(x, i + (w * 7) - 7)
                    .Update
                Next
            Next
        Next
    End With
    Set RstAgreement = Nothing
    Set RstSlots = Nothing
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
' Method DisplayAA()
' Displays Agreement on screen
'---------------------------------------------------------------
Public Sub DisplayAA()
    Dim RngCrewNo As Range
    Dim RngCrewName As Range
    Dim RngRole As Range
    Dim RngTemplateDate As Range
    Dim RngContractType As Range
    Dim RngHrsWk As Range
    Dim RngNoWeeks As Range
    Dim RngRevDate As Range
    Dim RngStation1 As Range
    Dim RngStation2 As Range
    Dim RngWeek1 As Range
    Dim i As Integer
    Dim x As Integer
    Dim w As Integer
    
    Dim WeekArry(1 To 48, 1 To 7)
    
    Set RngCrewNo = ShtFrontPage.Range("AY4")
    Set RngCrewName = ShtFrontPage.Range("AY5")
    Set RngRole = ShtFrontPage.Range("AY6")
    Set RngTemplateDate = ShtFrontPage.Range("AY7")
    Set RngContractType = ShtFrontPage.Range("AY8")
    Set RngHrsWk = ShtFrontPage.Range("AY9")
    Set RngNoWeeks = ShtFrontPage.Range("AY10")
    Set RngRevDate = ShtFrontPage.Range("AY11")
    Set RngStation1 = ShtFrontPage.Range("AY12")
    Set RngStation2 = ShtFrontPage.Range("AY13")
    Set RngWeek1 = ShtFrontPage.Range("B4:AW10")
    
    ClearAA
    
    RngCrewNo = pCrewMember.CrewNo
    RngCrewName = pCrewMember.Name
    RngRole = pCrewMember.Role
    RngTemplateDate = pAADate
    RngContractType = pContractType
    RngHrsWk = pHrsPW
    RngNoWeeks = pNoWeeks
    RngStation1 = pStation1.StationCallSign & " - " & pStation1.StationName
    RngStation2 = pStation2.StationCallSign & " - " & pStation2.StationName
    RngRevDate = pRevDateDue
        
    For w = 1 To pNoWeeks
        For i = 1 To 7
            For x = 1 To 48
                WeekArry(x, i) = pTimeSlot(x, i + (w * 7) - 7)
            Next
        Next
        RngWeek1.Offset((w * 8) - 8, 0) = Application.Transpose(WeekArry)
    Next
        
    Set RngCrewNo = Nothing
    Set RngCrewName = Nothing
    Set RngRole = Nothing
    Set RngTemplateDate = Nothing
    Set RngContractType = Nothing
    Set RngHrsWk = Nothing
    Set RngNoWeeks = Nothing
    Set RngRevDate = Nothing
    Set RngStation1 = Nothing
    Set RngStation2 = Nothing
    Set RngWeek1 = Nothing
End Sub

' ===============================================================
' Method ClearAA()
' Clears the displayed AA
'---------------------------------------------------------------
Public Sub ClearAA()
    ShtFrontPage.Range("B2:AW100").ClearContents
End Sub

' ===============================================================
' Method UpdateAA()
' Updates AA from screen
'---------------------------------------------------------------
Public Sub UpdateAA()
    Dim RngCrewNo As Range
    Dim RngCrewName As Range
    Dim RngRole As Range
    Dim RngTemplateDate As Range
    Dim RngContractType As Range
    Dim RngHrsWk As Range
    Dim RngNoWeeks As Range
    Dim RngRevDate As Range
    Dim RngWeek1 As Range
    Dim RngWeek2 As Range
    Dim i As Integer
    Dim x As Integer
    Dim w As Integer

    Dim WeekArry(1 To 48, 1 To 7)
    
    Set RngCrewNo = ShtFrontPage.Range("B2")
    Set RngCrewName = ShtFrontPage.Range("C2")
    Set RngRole = ShtFrontPage.Range("D2")
    Set RngTemplateDate = ShtFrontPage.Range("E2")
    Set RngContractType = ShtFrontPage.Range("F2")
    Set RngHrsWk = ShtFrontPage.Range("G2")
    Set RngNoWeeks = ShtFrontPage.Range("H2")
    Set RngRevDate = ShtFrontPage.Range("I2")
    Set RngWeek1 = ShtFrontPage.Range("B4:AW10")

    pCrewMember.Name = RngCrewName
    pCrewMember.Role = RngRole
    pAADate = RngTemplateDate
    pContractType = RngContractType
    pHrsPW = RngHrsWk
    pNoWeeks = RngNoWeeks
    pRevDateDue = RngRevDate
    
    For w = 1 To pNoWeeks
        For i = 1 To 7
            For x = 1 To 48
                pTimeSlot(x, i + (w * 7) - 7) = RngWeek1.Offset((w * 8) - 8, 0).Cells(i, x)
            Next
        Next
    Next
    
    
    
    Set RngCrewNo = Nothing
    Set RngCrewName = Nothing
    Set RngRole = Nothing
    Set RngTemplateDate = Nothing
    Set RngContractType = Nothing
    Set RngHrsWk = Nothing
    Set RngNoWeeks = Nothing
    Set RngRevDate = Nothing
    Set RngWeek1 = Nothing

End Sub
