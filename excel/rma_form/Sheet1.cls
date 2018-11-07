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
        Select Case ParseInputs(Trim(tRMANumber.Text))
            Case Is = 1
                SetFilter 1, tRMANumber.Text
            
            Case Is = 2
                
            Case Is = 3
                SetFilter 16, tRMANumber.Text
                
            Case Else
                SetFilter 0, ""

        End Select
        
    ElseIf tRMANumber.Text = "" Then
        
        wsDB.Range("A4:AA4").Columns.Hidden = False
        SetFilter 0, ""
    
    End If
    
End Sub


Private Sub tRMANumber_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        Select Case nStatus
            Case Is = 1
                wsDB.Range("C4:L4").Columns.Hidden = True
                nStatus = 11
            Case Is = 11
                wsDB.Range("C4:S4").Columns.Hidden = True
                nStatus = 12
            Case Is = 12
                wsDB.Range("C4:L4").Columns.Hidden = False
                wsDB.Range("M4:Y4").Columns.Hidden = True
                nStatus = 13
            Case Is = 13
                wsDB.Range("A4:AA4").Columns.Hidden = False
                nStatus = 1
                
            Case Is = 2
                Dim listRMAItems As IXMLDOMNodeList
                Dim nodeRMA As IXMLDOMNode
                Dim nNewRMANumber, nNewRMARow, nNewRMAOffset As Integer
                Dim i As Integer
                
                nNewRMANumber = wsDB.Range("CellNewRMANumber").Value
                nNewRMARow = wsDB.Range("CellNewRMARow").Value
                                
                If Not xmlRMA.LoadXML(tRMANumber.Text) Then
                    Err.Raise xmlRMA.parseError.ErrorCode, , xmlRMA.parseError.reason
                End If

                Set listRMAItems = xmlRMA.SelectNodes("//RMA/RMAItem")
                nNewRMAOffset = listRMAItems.Length - 1
                
                wsDB.Range("C" & nNewRMARow & ":C" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/Company").nodeTypedValue
                wsDB.Range("D" & nNewRMARow & ":D" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/ContactPerson").nodeTypedValue
                wsDB.Range("E" & nNewRMARow & ":E" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/Address1").nodeTypedValue
                wsDB.Range("F" & nNewRMARow & ":F" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/Address2").nodeTypedValue
                wsDB.Range("G" & nNewRMARow & ":G" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/City").nodeTypedValue
                wsDB.Range("H" & nNewRMARow & ":H" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/PostCode").nodeTypedValue
                wsDB.Range("I" & nNewRMARow & ":I" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/PhoneNumber").nodeTypedValue
                wsDB.Range("J" & nNewRMARow & ":J" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/Email").nodeTypedValue
                wsDB.Range("K" & nNewRMARow & ":K" & nNewRMARow + nNewRMAOffset).Value = xmlRMA.SelectSingleNode("//RMA/SpecialInstruction").nodeTypedValue
                wsDB.Range("L" & nNewRMARow & ":L" & nNewRMARow + nNewRMAOffset).Value = Date
                
                i = 0
                For Each nodeRMA In listRMAItems
                    wsDB.Range("A" & nNewRMARow + i).Value = "RMA" & nNewRMANumber & "-" & i
                    wsDB.Range("B" & nNewRMARow + i).Value = "RMA Requested Awaiting Arrival"
                    wsDB.Range("P" & nNewRMARow + i).Value = nodeRMA.SelectSingleNode("SerialNumber").nodeTypedValue
                    wsDB.Range("T" & nNewRMARow + i).Value = nodeRMA.SelectSingleNode("FaultDescription").nodeTypedValue
                    
                    i = i + 1
                Next nodeRMA
                
                tRMANumber.Text = ""

                wsDB.Range("CellNewRMANumber").Value = nNewRMANumber + 1
                wsDB.Range("CellNewRMARow").Value = nNewRMARow + i

            Case Is = 3
                wsDB.Range("D4:K4").Columns.Hidden = True

            Case Else
                
        End Select
        
    End If
End Sub


