Attribute VB_Name = "Test"

Public Sub GetAA()
    Dim AA As ClsAgreement
    
    Set AA = New ClsAgreement
    Initialise
    AA.CrewNo = "1015"
    AA.DBGet
    AA.DisplayAA
    Set AA = Nothing
End Sub

Public Sub UpdateAA()
    Dim AA As ClsAgreement
    Initialise
    Set AA = New ClsAgreement
    AA.CrewNo = "1015"
    AA.DBGet
'    AA.DisplayAA
    AA.UpdateAA
    AA.DBSave
    Set AA = Nothing
End Sub


Public Sub TestStnLookUp()
    Dim Stn As TypeStation
    
    Stn = ModDBLookups.StationLookUp(StationName:="Lincoln North")
    
    Debug.Print Stn.StationCallSign, Stn.StationName, Stn.StationNo
End Sub
