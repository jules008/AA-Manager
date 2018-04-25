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
    Dim RstTemplate As Recordset
    Dim RstTemplateDetail As Recordset
    Dim RstTemplateStns As Recordset
    
    Set RstTemplate = ModDatabase.SQLQuery("SELECT * FROM TblTemplate WHERE CrewNo = '" & pCrewMember.CrewNo & "'")
    Set RstTemplateDetail = ModDatabase.SQLQuery("SELECT * FROM TblTemplateDetail WHERE CrewNo = '" & pCrewMember.CrewNo & "'")
    Set RstTemplateStns = ModDatabase.SQLQuery("SELECT * FROM TblTemplateStns WHERE CrewNo = '" & pCrewMember.CrewNo & "'")
   
    With RstTemplate
        If .RecordCount > 0 Then
            If Not IsNull(!CrewNo) Then pCrewMember.CrewNo = !CrewNo
            If Not IsNull(!Role) Then pCrewMember.Role = !Role
            If Not IsNull(!CrewName) Then pCrewMember.Name = !CrewName
            If Not IsNull(!TemplateDate) Then pAADate = !TemplateDate
            If Not IsNull(!ContractType) Then pContractType = !ContractType
            If Not IsNull(!HrsPW) Then pHrsPW = !HrsPW
            If Not IsNull(!NoWeeks) Then pNoWeeks = !NoWeeks
            If Not IsNull(!RevDateDue) Then pRevDateDue = !RevDateDue
        End If
    End With
        
    With RstTemplateStns
        Do While Not .EOF
            If !Station = 1 Then
                If Not IsNull(!StationNo) Then pStation1 = ModDBLookups.StationLookUp(StationNo:=!StationNo)
            Else
                If Not IsNull(!StationNo) Then pStation2 = ModDBLookups.StationLookUp(StationNo:=!StationNo)
            End If
            .MoveNext
        Loop
    End With
            
    With RstTemplateDetail
        If .RecordCount > 0 Then
            Do While Not .EOF
                pTimeSlot(!TimeSlot, !Day) = !OnCall
                .MoveNext
            Loop
        End If
    End With
    Set RstTemplate = Nothing
    Set RstTemplateDetail = Nothing
    Set RstTemplateStns = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstTemplate As Recordset
    Dim RstTemplateDetail As Recordset
    Dim RstTemplateStns As Recordset
    Dim i As Integer
    Dim x As Integer
    Dim w As Integer

    Set RstTemplate = ModDatabase.SQLQuery("SELECT * FROM TblTemplate WHERE CrewNo = '" & pCrewMember.CrewNo & "'")
    Set RstTemplateDetail = ModDatabase.SQLQuery("TblTemplateDetail")
    Set RstTemplateStns = ModDatabase.SQLQuery("TblTemplateStns")
   
    With RstTemplate
        If .RecordCount > 0 Then
            .Edit
            !CrewNo = pCrewMember.CrewNo
            !CrewName = pCrewMember.Name
            !Role = pCrewMember.Role
            If pAADate <> 0 Then !TemplateDate = pAADate
            !ContractType = pContractType
            !HrsPW = pHrsPW
            !NoWeeks = pNoWeeks
            !RevDateDue = pRevDateDue
            .Update
        End If
    End With
            
    DB.Execute ("DELETE FROM TblTemplateStns WHERE CrewNo = '" & pCrewMember.CrewNo & "'")
    
    With RstTemplateStns
        If pStation1.StationNo <> 0 Then
            .AddNew
            !CrewNo = pCrewMember.CrewNo
            !Station = 1
            !StationNo = pStation1.StationNo
            .Update
        End If
        
        If pStation2.StationNo <> 0 Then
            .AddNew
            !CrewNo = pCrewMember.CrewNo
            !Station = 2
            !StationNo = pStation2.StationNo
            .Update
        End If
    End With
    
    DB.Execute ("DELETE FROM TblTemplateDetail WHERE CrewNo = '" & pCrewMember.CrewNo & "'")
    
    With RstTemplateDetail
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
    Set RstTemplate = Nothing
    Set RstTemplateDetail = Nothing
    Set RstTemplateStns = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstTemplate As Recordset
    Dim i As Integer

    Set RstTemplate = ModDatabase.SQLQuery("SELECT * FROM TblAgreement WHERE CrewNo = " & pCrewNo & " AND Deleted IS NULL")
    With RstTemplate
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

    Set RstTemplate = Nothing
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
    RngContractType = ModDBLookups.ContractLookUp(pContractType)
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
    Dim RngStation1 As Range
    Dim RngStation2 As Range
    Dim RngWeek1 As Range
    Dim StationStr() As String
    Dim i As Integer
    Dim x As Integer
    Dim w As Integer

    Dim WeekArry(1 To 48, 1 To 7)
    
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

    pCrewMember.Name = RngCrewName
    pCrewMember.Role = RngRole
    pAADate = RngTemplateDate
    If RngContractType = "Under 120hrs" Then pContractType = 1 Else pContractType = 2
    pHrsPW = RngHrsWk
    pNoWeeks = RngNoWeeks
    pRevDateDue = RngRevDate
    
    If RngStation1 <> "" Then
        StationStr = Split(RngStation1, " - ")
        pStation1 = ModDBLookups.StationLookUp(Callsign:=StationStr(0))
    End If
    
    If RngStation2 <> "" Then
        StationStr = Split(RngStation2, " - ")
        pStation2 = ModDBLookups.StationLookUp(Callsign:=StationStr(0))
    End If
    
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
    Set RngStation1 = Nothing
    Set RngStation2 = Nothing
    Set RngWeek1 = Nothing

End Sub