Attribute VB_Name = "DashBoard"
Public wbRMA As Workbook
Public wsDB As Worksheet
Public tblRMA As ListObject
Public xmlRMA As MSXML2.DOMDocument
Public bInitialized As Boolean
Public nStatus As Integer
Public nFilter As Integer
Public regX As RegExp

Sub Initialize()
    
    Set wbRMA = ActiveWorkbook
    Set wsDB = wbRMA.Sheets(1)
    Set tblRMA = wsDB.ListObjects("tableRMA")
    Set xmlRMA = New MSXML2.DOMDocument
    Set regX = New RegExp
    
    'rmanumber text box get focused while the excel book opened.
    ActiveSheet.tRMANumber.Activate
    ActiveSheet.tRMANumber.Text = ""

    tblRMA.AutoFilter.ShowAllData
    
    nStatus = 0

    bInitialized = True
    
End Sub

Function ParseInputs(txIn) As Integer
    
    If UCase(Left(txIn, 3)) = "RMA" And nStatus < 5 Then    'less than 5 = no commands mode anymore
        nStatus = 1
        
        If Len(txIn) = 13 And Right(txIn, 1) = " " Then     'entering command mode, setup right filter and columns
            nStatus = 5
        End If
        
    ElseIf Left(txIn, 1) = "<" Then
        nStatus = 2
    
    ElseIf Len(txIn) > 12 And nStatus >= 5 Then    'part number command mode to update part number and status
        nStatus = 6
            
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
        Case Is = 0
            tblRMA.AutoFilter.ShowAllData
            
        Case Is = 1
            tblRMA.Range.AutoFilter Field:=1, Criteria1:="=" & txIn & "*"
            
        Case Is = 16
            tblRMA.Range.AutoFilter Field:=16, Criteria1:="=*" & txIn & "*"
            
    End Select
    
End Sub

Sub ResetView()
    tblRMA.Range.Columns.Hidden = False
    tblRMA.AutoFilter.ShowAllData
End Sub
