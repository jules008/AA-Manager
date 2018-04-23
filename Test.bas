Attribute VB_Name = "Test"

Public Sub GetAA()
    Dim AA As ClsAgreement
    
    Set AA = New ClsAgreement
    
    AA.CrewNo = "5398"
    AA.DBGet
    AA.DisplayAA
    Set AA = Nothing
End Sub
