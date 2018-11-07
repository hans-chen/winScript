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
    
    'rmanumber text box get focused while the excel book opened.
    ActiveSheet.tRMANumber.Activate
    ActiveSheet.tRMANumber.Text = ""

    tblRMA.AutoFilter.ShowAllData
    
    nStatus = 0

    bInitialized = True
    
End Sub

Function ParseInputs(txIn) As Integer
    
    If UCase(Left(txIn, 3)) = "RMA" Then
        nStatus = 1

        If Len(txIn) = 14 & UCase(Left(txIn, 2)) = "PN" Then
            nStatus = 4
        End If

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
