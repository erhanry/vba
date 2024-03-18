Attribute VB_Name = "Remit"

Sub uploadxml()

  Dim FileDialog As Office.FileDialog
  Dim xmlDoc As Object

  Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("List")
  Dim conf As Worksheet: Set conf = ThisWorkbook.Sheets("Config")
  Dim reportingEntityID As Variant: Set reportingEntityID = conf.ListObjects("reportingEntityID").DataBodyRange
  Dim contractList As Variant: Set contractList = conf.ListObjects("contractList").DataBodyRange
  Dim OrderList As Variant: Set OrderList = conf.ListObjects("OrderList").DataBodyRange
  Dim TradeList As Variant: Set TradeList = conf.ListObjects("TradeList").DataBodyRange

Application.ScreenUpdating = False
Application.DisplayAlerts = False

  Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
  With FileDialog
    .Filters.Clear
    .Title = "Select a Remit XML File"
    .Filters.Add "Remit XML File", "*.xml", 1
    .AllowMultiSelect = False

    If .Show = True Then
      xmlFileName = .SelectedItems(1)
      Set xmlDoc = CreateObject("MSXML2.DOMDocument")
      xmlDoc.async = False: xmlDoc.validateOnParse = False
      xmlDoc.Load (xmlFileName)
   If (xmlDoc.parseError.ErrorCode <> 0) Then
   Dim myErr
   Set myErr = xmlDoc.parseError
   MsgBox ("You have error " & myErr.reason)
End If

'----------------------------------- OrderList -------------------------------------
For Each Segment In xmlDoc.SelectNodes("//OrderList/OrderReport")
  i = LastRow

  ws.Range("A" & i).Value = "OrderReport"

  If reportingEntityID(1, 2) <> "" Then
    ws.Range(reportingEntityID(1, 2) & i).Value = xmlDoc.SelectSingleNode(reportingEntityID(1, 1)).Text
  End If
  
  For Item = 1 To OrderList.Rows.Count
    If OrderList(Item, 2) <> "" Then
      ws.Range(OrderList(Item, 2) & i).Value = Segment.SelectSingleNode(OrderList(Item, 1)).Text
    End If
  Next Item
  contractId = Segment.SelectSingleNode("contractInfo/contractId").Text

For Each Key In xmlDoc.SelectNodes("//contractList/contract")
  For Item1 = 1 To contractList.Rows.Count
    If contractList(Item1, 2) <> "" And Key.SelectSingleNode("contractId").Text = contractId Then
      ws.Range(contractList(Item1, 2) & i).Value = Key.SelectSingleNode(contractList(Item1, 1)).Text
    End If
  Next Item1
Next Key
Next Segment

'---------------------------------- TradeReport ------------------------------------
For Each Segment In xmlDoc.SelectNodes("//TradeList/TradeReport")
  i = LastRow

  ws.Range("A" & i).Value = "TradeReport"

  If reportingEntityID(1, 2) <> "" Then
    ws.Range(reportingEntityID(1, 2) & i).Value = xmlDoc.SelectSingleNode(reportingEntityID(1, 1)).Text
  End If
  
  For Item = 1 To TradeList.Rows.Count
    If TradeList(Item, 2) <> "" Then
      ws.Range(TradeList(Item, 2) & i).Value = Segment.SelectSingleNode(TradeList(Item, 1)).Text
    End If
  Next Item
  da = Segment.SelectSingleNode("contractInfo/contractId").Text

For Each ara In xmlDoc.SelectNodes("//contractList/contract")
  For Item1 = 1 To contractList.Rows.Count
    If contractList(Item1, 2) <> "" And ara.SelectSingleNode("contractId").Text = da Then
      ws.Range(contractList(Item1, 2) & i).Value = ara.SelectSingleNode(contractList(Item1, 1)).Text
    End If
  Next Item1
Next ara

Next Segment
'-----------------------------------------------------------------------------------

     End If
  End With
If LastRow > 30 Then
 Application.Goto ws.Cells(LastRow - 30, "A"), True
 End If

End Sub
Function LastRow()
    LastRow = ThisWorkbook.Sheets("List").Cells(1, "A").End(xlDown).Row + 1
  If LastRow > 1048500 Then
    LastRow = ThisWorkbook.Sheets("List").Cells(1, "A").End(xlUp).Row + 1
  End If
End Function
