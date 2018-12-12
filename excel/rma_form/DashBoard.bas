Attribute VB_Name = "DashBoard"
Public wbRMA As Workbook
Public wsDB As Worksheet
Public tblRMA As ListObject
Public xmlRMA As MSXML2.DOMDocument
Public bInitialized As Boolean
Public nStatus As Integer
Public nFilter As Integer
Public regX As RegExp
Public nRMALastNumber


Sub Initialize()
    
    Set wbRMA = ActiveWorkbook
    Set wsDB = wbRMA.Sheets(1)
    Set tblRMA = wsDB.ListObjects("tableRMA")
    Set xmlRMA = New MSXML2.DOMDocument
    Set regX = New RegExp
    
    Worksheets("Dashboard").Activate
    
    'rmanumber text box get focused while the excel book opened.
    ActiveSheet.tRMANumber.Activate
    ActiveSheet.tRMANumber.Text = ""

    tblRMA.AutoFilter.ShowAllData
    
    nStatus = 0

    bInitialized = True
    
End Sub

Function ParseInputs(txIn) As Integer
    
    If nStatus >= 5 Then
    
    ElseIf Left(txIn, 1) = "<" Then
        nStatus = 2

    ElseIf Right(txIn, 1) = " " Then     'entering command mode, setup right filter and columns
       nStatus = 5

    ElseIf UCase(Left(txIn, 3)) = "RMA" And nStatus < 5 Then    'less than 5 = no commands mode anymore
        nStatus = 1
      
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

Sub SendMailMessage(txNumber, txEmailAddr)
    Dim OutApp As Outlook.Application
    Dim objOutlookMsg As Outlook.MailItem
    Dim objOutlookRecip As Recipient
    Dim Recipients As Recipients
    Dim strFilename As String: strFilename = "\\freebsd\guest\email.html"
    Dim strFilecontent, strEmail, strProducts, strSNs As String
    Dim iFile As Integer: iFile = FreeFile
    Dim nCount As Integer
    
    Open strFilename For Input As #iFile
    strFilecontent = Input(LOF(iFile), iFile)
    Close #iFile
    
    Set OutApp = CreateObject("Outlook.Application")
    Set objOutlookMsg = OutApp.CreateItem(olMailItem)
    
    Set Recipients = objOutlookMsg.Recipients
    Set objOutlookRecip = Recipients.Add(txEmailAddr)
    objOutlookRecip.Type = 1
    
    nCount = tblRMA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Rows.Count
    
    If nCount < 1 Then
        Exit Sub
    End If
    
    i = 1
    strHTMLTableRow = "<tr><td valign=""top"" class=""mcnTextContent"" style=""padding-top: 0;padding-left: 18px;padding-bottom: 9px;" & _
                    "padding-right: 18px;mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;word-break: break-word;color: #202020;" & _
                    "font-family: 'Lato', 'Helvetica Neue', Helvetica, Arial, sans-serif;font-size: 16px;line-height: 150%;text-align: left;"">"
    
    Do While i <= nCount
        strProducts = strProducts & strHTMLTableRow & tblRMA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(i, 16).Value & "</td></tr>"
        strSNs = strSNs & strHTMLTableRow & tblRMA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(i, 17).Value & "</td></tr>"
        
        i = i + 1
    Loop
    
    strEmail = Replace(strFilecontent, "+++Contact+++", tblRMA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 4).Value)
    strEmail = Replace(strEmail, "+++RMANUMBER+++", txNumber)
    strEmail = Replace(strEmail, "+++RMAPRODUCTS+++", strProducts)
    strEmail = Replace(strEmail, "+++RMASNS+++", strSNs)
    
    With objOutlookMsg
        '.SentOnBehalfOfName = "Rick.Cranen@newland-id.com"
        .Subject = "RMA Number: " & txNumber
        .Attachments.Add "\\freebsd\guest\pics\newlandlogobanner.jpg", olByValue, 0
        .Attachments.Add "\\freebsd\guest\pics\linkedinbanner.jpg", olByValue, 0
        .HTMLBody = strEmail

        'Resolve each Recipient's name.
        For Each objOutlookRecip In objOutlookMsg.Recipients
            objOutlookRecip.Resolve
        Next
        '.ReplyRecipients.Add "Rick.Cranen@newland-id.com"
        '.Sender.Address = "Rick.Cranen@newland-id.com"
        .display
    End With
    
      'objOutlookMsg.Send
    Set OutApp = Nothing

End Sub

