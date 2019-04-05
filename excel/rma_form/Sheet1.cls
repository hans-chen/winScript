VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub tRMANumber_Change()
    If Len(tRMANumber.Text) > 3 Then
        Select Case ParseInputs(tRMANumber.Text)
            Case Is = 1
                SetFilter 1, tRMANumber.Text
            
            Case Is = 2
                
            Case Is = 3
                SetFilter 16, tRMANumber.Text
                
            Case Is = 5
                nStatus = 6
                tRMANumber.Text = Trim(tRMANumber.Text) & ">"   'change space to >
                
            Case Is = 6     'dont change filter or hidden columns
            Case Is = 99
            
            Case Else
                SetFilter 0, ""

        End Select
        
    ElseIf tRMANumber.Text = "" Then
        nStatus = 0
        ResetView
        
    ElseIf tRMANumber.Text = " " Then
        nStatus = 6
        tRMANumber.Text = ">"
    
    End If
    
End Sub


Private Sub tRMANumber_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 38 Then
        Select Case nStatus
        Case Is = 0
            nRMALastNumber = wsDB.Range("CellNewRMANumber").Value2 - 1
            tRMANumber.Text = "RMA" & nRMALastNumber
        
        Case Is = 1
            n = CLng(Mid(tRMANumber.Text, 4, 6))
            tRMANumber.Text = "RMA" & n - 1
            
        End Select
    End If
    
    If KeyCode = 40 Then
        Select Case nStatus
        Case Is = 0
            nRMALastNumber = wsDB.Range("CellNewRMANumber").Value2 + 1
            tRMANumber.Text = "RMA" & nRMALastNumber
        
        Case Is = 1
            n = CLng(Mid(tRMANumber.Text, 4, 6))
            tRMANumber.Text = "RMA" & n + 1
            
        End Select
    End If
    
    If KeyCode = 13 Then
        Select Case nStatus
        Case Is = 1                                     'switch to RMA received updating mode, hide all customer's columns
            wsDB.Range("C4:M4").Columns.Hidden = True
            nStatus = 11
        Case Is = 11                                    'switch to RMA repaired updating mode, hide all cusotmer's and product's columns
            wsDB.Range("C4:T4").Columns.Hidden = True
            nStatus = 12
        Case Is = 12                                    'switch to RMA shipped back updating mode, hide all repairing's and product's columns
            wsDB.Range("C4:M4").Columns.Hidden = False
            wsDB.Range("M4:Z4").Columns.Hidden = True
            nStatus = 13
        Case Is = 13                                    'show all columns and reset to RMA number working mode
            wsDB.Range("A4:AB4").Columns.Hidden = False
            nStatus = 1
            
        Case Is = 2                                     'xml working mode to add new RMA items
            Dim listRMAItems As IXMLDOMNodeList
            Dim nodeRMA As IXMLDOMNode
            Dim nNewRMANumber, nNewRMARow, nNewRMAOffset As Integer
            Dim i As Integer
            Dim strRawXML As String
                        
            nNewRMANumber = wsDB.Range("CellNewRMANumber").Value
            nNewRMARow = wsDB.Range("CellNewRMARow").Value
            strRawXML = Replace(tRMANumber.Text, "&", "and")
                            
            If Not xmlRMA.LoadXML(strRawXML) Then
                Err.Raise xmlRMA.parseError.ErrorCode, , xmlRMA.parseError.reason
            End If

            Set listRMAItems = xmlRMA.SelectNodes("//RMA/RMAItems/RMAItem")
            nNewRMAOffset = listRMAItems.Length - 1
            
            'add multiple lines common information of one rma request
            wsDB.Range("C" & nNewRMARow & ":C" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/Company").nodeTypedValue
            wsDB.Range("D" & nNewRMARow & ":D" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/ContactPerson").nodeTypedValue
            wsDB.Range("E" & nNewRMARow & ":E" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/Address1").nodeTypedValue
            wsDB.Range("F" & nNewRMARow & ":F" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/Address2").nodeTypedValue
            wsDB.Range("G" & nNewRMARow & ":G" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/City").nodeTypedValue
            wsDB.Range("H" & nNewRMARow & ":H" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/PostCode").nodeTypedValue
            wsDB.Range("I" & nNewRMARow & ":I" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/Country").nodeTypedValue
            wsDB.Range("J" & nNewRMARow & ":J" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/PhoneNumber").nodeTypedValue
            wsDB.Range("K" & nNewRMARow & ":K" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/Email").nodeTypedValue
            wsDB.Range("L" & nNewRMARow & ":L" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/SpecialInstruction").nodeTypedValue
            wsDB.Range("M" & nNewRMARow & ":M" & nNewRMARow + nNewRMAOffset).Value = Date
            
            'add rma request information per rma item
            i = 0
            For Each nodeRMA In listRMAItems
                wsDB.Range("A" & nNewRMARow + i).Value = "RMA" & nNewRMANumber & "-" & WorksheetFunction.Text(i, "00")
                wsDB.Range("B" & nNewRMARow + i).Value = "RMA Requested Awaiting Arrival"
                wsDB.Range("P" & nNewRMARow + i).Value = nodeRMA.SelectSingleNode("Product").nodeTypedValue
                wsDB.Range("Q" & nNewRMARow + i).Value = nodeRMA.SelectSingleNode("SerialNumber").nodeTypedValue
                wsDB.Range("U" & nNewRMARow + i).Value = nodeRMA.SelectSingleNode("FaultDescription").nodeTypedValue
                
                i = i + 1
            Next nodeRMA
            
            'update the cell has last rma number and last rma record row number
            wsDB.Range("CellNewRMANumber").Value = nNewRMANumber + 1
            wsDB.Range("CellNewRMARow").Value = nNewRMARow + i
            
            tRMANumber.Text = ""

        Case Is = 3                                     'serial number working mode
            wsDB.Range("D4:K4").Columns.Hidden = True
            
        Case Is = 6
            Dim mcCmd As MatchCollection
            Dim mCmd As Match
            
            strcmd = tRMANumber.Text
            With regX
                .Global = True
                .IgnoreCase = True
                .Pattern = "([^>| ]+)"
            End With
            
            Set mcCmd = regX.Execute(strcmd)
            
            For Each mCmd In mcCmd
                Select Case UCase(mCmd.Value)
                
                Case Is = "PN"
                   tblRMA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 15).Value = UCase(Mid(mCmd.Value, 3))
                   
                Case Is = "WR"
                   tblRMA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 17).Value = UCase(Mid(mCmd.Value, 3))
                   
                Case Is = "REPORT"
                    wbRMA.Sheets("CSV").Copy
                    If Dir("\\freebsd\guest\rmareport.csv") <> "" Then
                        Kill "\\freebsd\guest\rmareport.csv"
                    End If
                    ActiveWorkbook.SaveAs Filename:="\\freebsd\guest\rmareport.csv", FileFormat:=xlCSVMSDOS, CreateBackup:=False
                    ActiveWorkbook.Close SaveChanges:=False
                    
                Case Is = "EMAIL"
                    tNumber = mcCmd.Item(0).Value                                                                               'get rma number from command
                    tEmailReciept = tblRMA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 11).Value        'get the email address from filtered data
    
                    SendMailMessage tNumber, tEmailReciept, "RECEIPT"
                    
                Case Is = "COST"
                    tNumber = mcCmd.Item(0).Value                                                                               'get rma number from command
                    tEmailReciept = tblRMA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 11).Value        'get the email address from filtered data
    
                    SendMailMessage tNumber, tEmailReciept, "COST"
                        
                End Select
            Next

        Case Else

        End Select
        
    End If
End Sub



