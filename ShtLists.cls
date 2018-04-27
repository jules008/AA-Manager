VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Sheet ShtLists
' v0,0 - Initial Version
' v0,1 - Added keyword search
' v0,2 - Don't display hidden assets
' v0,3 - Bug Fix, clear asset list before reloading
'---------------------------------------------------------------
' Date - 06 Nov 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ShtLists"

' ===============================================================
' RefreshNameList
' Updates the list of assets from the database
' ---------------------------------------------------------------
Public Function RefreshNameList() As Boolean
    Dim RstNames As Recordset
    
    Const StrPROCEDURE As String = "RefreshNameList()"

    On Error GoTo ErrorHandler

    Range("A1:B3000").ClearContents
    
    Set RstNames = ModDatabase.SQLQuery("SELECT CrewNo, CrewName FROM TblTemplate ORDER BY CrewNo")

    Me.Range("A1").CopyFromRecordset RstNames

    RefreshNameList = True
    
    Set RstNames = Nothing

Exit Function

ErrorExit:
    
    Terminate

    Set RstNames = Nothing
    
    RefreshNameList = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function







