Attribute VB_Name = "DashBoard"
Public wbRMA As Workbook
Public wsDB As Worksheet
Public tblRMA As ListObject
Public xmlRMA As MSXML2.DOMDocument
Public bInitialized As Boolean
Public nStatus As Integer
Public nFilter As Integer

Sub Initialize()
    
    Set wbRMA = ActiveWorkbook
    Set wsDB = wbRMA.Sheets(1)
    Set tblRMA = wsDB.ListObjects("tableRMA")
    Set xmlRMA = New MSXML2.DOMDocument
    
    ActiveSheet.tRMANumber.Activate
    ActiveSheet.tRMANumber.Text = ""

    tblRMA.AutoFilter.ShowAllData
    
    bInitialized = True
    
End Sub

Function ParseInputs(txIn) As Integer

    nStatus = 0
    
    If UCase(Left(txIn, 3)) = "RMA" Then
        nStatus = 1
    ElseIf Left(txIn, 1) = "<" Then
        nStatus = 2
    ElseIf Len(txIn) > 0 Then
        nStatus = 3
    End If
    
    ParseInputs = nStatus

End Function

Sub SetFilter(nFilterColumn, txIn)
    If nFilter <> nFilterColumn Then
        tblRMA.AutoFilter.ShowAllData
        nFilter = nFilterColumn
        
    End If
    
    Select Case nFilterColumn
        Case Is = 1
            tblRMA.Range.AutoFilter Field:=1, Criteria1:="=" & txIn & "*"
            
        Case Is = 16
            tblRMA.Range.AutoFilter Field:=16, Criteria1:="=*" & txIn & "*"
            
    End Select
    
End Sub
