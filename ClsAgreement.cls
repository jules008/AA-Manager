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
' Date - 22 Oct 18
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
Private pContractType As EnumContractType
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
Public Property Get ContractType() As EnumContractType
    ContractType = pContractType
End Property

Public Property Let ContractType(ByVal vNewValue As EnumContractType)
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
        If .RecordCount = 0 Then
            .AddNew
            !CrewNo = pCrewMember.CrewNo
        Else
            .Edit
        End If
        
        !CrewName = pCrewMember.Name
        !Role = pCrewMember.Role
        If pAADate <> 0 Then !TemplateDate = pAADate
        !ContractType = pContractType
        !HrsPW = pHrsPW
        !NoWeeks = pNoWeeks
        !RevDateDue = pRevDateDue
        .Update
        
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
' Method Display()
' Displays Agreement on screen
'---------------------------------------------------------------
Public Sub Display()
    Dim RngWeek1 As Range
    Dim i As Integer
    Dim x As Integer
    Dim w As Integer
    Dim AccessGranted As EnumTriState
    
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Dim WeekArry(1 To 48, 1 To 7)
    
    If CurrentUser Is Nothing Then Exit Sub
    
    AccessGranted = ModSecurity.StationAccessCheck(pStation1.StationNo, pStation2.StationNo)
    
    If AccessGranted = xTrue Then
        Set RngWeek1 = ShtFrontPage.Range(RNG_WEEK_1)
           
        With ShtFrontPage
            .TxtCrewNo = pCrewMember.CrewNo
            .TxtName = pCrewMember.Name
            .CmoRole.ListIndex = pCrewMember.Role
            If pAADate <> 0 Then .TxtTemplateDate = pAADate Else .TxtTemplateDate = ""
            .CmoContract.ListIndex = pContractType
            .CmoNoWeeks.ListIndex = pNoWeeks - 1
            .CmoStation1.ListIndex = pStation1.StationNo - 1
            .CmoStation2.ListIndex = pStation2.StationNo - 1
            
            If pRevDateDue <> 0 Then
                .TxtRevDate.Visible = True
                .TxtRevDate = Format(pRevDateDue, "dd mmm yy")
                If Not .FormatRevDate Then Err.Raise HANDLED_ERROR

            Else
                .TxtRevDate.Visible = False
                .TxtRevDate = ""
            End If
        End With
        
        If Not ShtFrontPage.FormatWkRnge(pNoWeeks) Then Err.Raise HANDLED_ERROR
        
        For w = 1 To pNoWeeks
            For i = 1 To 7
                For x = 1 To 48
                    WeekArry(x, i) = pTimeSlot(x, i + (w * 7) - 7)
                Next
            Next
            RngWeek1.Offset((w * 9) - 9, 0) = Application.Transpose(WeekArry)
        Next
    Else
        MsgBox "Sorry, you are not permitted to view Availability Agreements from this station.  Please see SM Andy Dixon for access", vbOKOnly + vbExclamation, APP_NAME
    End If
    
    If Not ShtFrontPage.UpdateValues Then Err.Raise HANDLED_ERROR
    
    Set RngWeek1 = Nothing
    
    Application.ScreenUpdating = True

Exit Sub

ErrorHandler:
    
    Application.ScreenUpdating = True

    Set RngWeek1 = Nothing
End Sub

' ===============================================================
' Method Clear()
' Clears the displayed AA
'---------------------------------------------------------------
Public Sub Clear()
    ShtFrontPage.Clear
End Sub

' ===============================================================
' Method Update()
' Updates AA from screen
'---------------------------------------------------------------
Public Sub Update()
    Dim RngWeek1 As Range
    Dim StationStr() As String
    Dim i As Integer
    Dim x As Integer
    Dim w As Integer

    Set RngWeek1 = ShtFrontPage.Range(RNG_WEEK_1)
    
    Dim WeekArry(1 To 48, 1 To 7)
    
    With ShtFrontPage
        pCrewMember.Name = .TxtName
        pCrewMember.Role = .CmoRole.ListIndex
        
        If .TxtTemplateDate <> "" Then pAADate = .TxtTemplateDate
        
        pContractType = .CmoContract.ListIndex
        pHrsPW = 0
        pNoWeeks = .CmoNoWeeks.ListIndex + 1
        
        If .TxtRevDate <> "" Then pRevDateDue = Format(.TxtRevDate, "dd mmm yy")
        
        If .CmoStation1.ListIndex <> -1 Then
            pStation1 = ModDBLookups.StationLookUp(StationNo:=.CmoStation1.ListIndex + 1)
        End If
    
        If .CmoStation2.ListIndex <> -1 Then
            pStation2 = ModDBLookups.StationLookUp(StationNo:=.CmoStation2.ListIndex + 1)
        End If
    End With
    
    For w = 1 To pNoWeeks
        For i = 1 To 7
            For x = 1 To 48
                pTimeSlot(x, i + (w * 7) - 7) = RngWeek1.Offset((w * 8) - 8, 0).Cells(i, x)
            Next
        Next
    Next
    
    Set RngWeek1 = Nothing

End Sub

' ===============================================================
' Method CreateNew()
' Creates new agreement
'---------------------------------------------------------------
Public Sub CreateNew()
    Dim CrewNo As String
    Dim Response As Integer
    Dim Continue As Boolean
    
    Do While Continue = False
        CrewNo = Application.InputBox("Please enter the Brigade No for the new Template", APP_NAME, Type:=7)
        
        If CrewNo = "False" Then Exit Sub
        
        pCrewMember = ModDBLookups.PersonLookUp(CrewNo)
        
        'check to see whether crewmember already has a template in the DB.  If so, offer the opportunity to open.
        'if they decline to open the template, restart the process so that a new brigade number can be added
        With pCrewMember
            If .CrewNo <> "" Then
                Response = MsgBox(.CrewNo & " - " & .Name & " already has a template saved.  Do you want to open this?  if you select No, " _
                                            & "you can enter another Brigade No", vbYesNoCancel + vbExclamation, APP_NAME)
                If Response = 2 Then Exit Sub
                
                If Response = 6 Then
                    DBGet
                    Display
                    Continue = True
                    Exit Sub
                Else
                    Continue = False
                End If
            Else
                pCrewMember.CrewNo = CrewNo
                Continue = True
            End If
        End With
    Loop
    
    pNoWeeks = Application.InputBox("How many weeks in the template cycle?", APP_NAME)
    
    pAADate = Format(Now, "dd/mm/yy")
    
    If Not ShtFrontPage.FormatWkRnge(pNoWeeks) Then Err.Raise HANDLED_ERROR
    
'    Display
    DBSave
    Display
End Sub
