Attribute VB_Name = "Remit"
Public Property Get ws() As Worksheet
  Set ws = ThisWorkbook.Worksheets("list")
End Property

Public Property Get conf() As Worksheet
  Set conf = ThisWorkbook.Worksheets("Config")
End Property

Function LastRow()
    LastRow = ws.Cells(1, "A").End(xlDown).Row + 1
  If LastRow > 1048500 Then
    LastRow = ws.Cells(1, "A").End(xlUp).Row + 1
  End If
End Function

Function node_check(ByRef Nodes As MSXML2.IXMLDOMNode)
  If Not Nodes Is Nothing Then node_check = Nodes.Text
End Function

Sub uploadxml()
Dim ST As Single: ST = Timer

Dim FileDialog As Office.FileDialog: Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
Dim SelectedFile As FileSystemObject: Set SelectedFile = New FileSystemObject
Dim xmlDoc As MSXML2.DOMDocument: Set xmlDoc = New DOMDocument

Dim reportingEntityID As Variant: Set reportingEntityID = conf.ListObjects("reportingEntityID").DataBodyRange
Dim contractList As Variant: Set contractList = conf.ListObjects("contractList").DataBodyRange
Dim OrderList As Variant: Set OrderList = conf.ListObjects("OrderList").DataBodyRange
Dim QuantityDetails As Variant: Set QuantityDetails = conf.ListObjects("priceIntervalQuantityDetails").DataBodyRange
Dim TradeList As Variant: Set TradeList = conf.ListObjects("TradeList").DataBodyRange
Dim Path, xmlFile, newFile, Title, contractId As String: Path = ThisWorkbook.Path & "\archive\": Title = "Импортиране на файл!"

Application.ScreenUpdating = False
Application.EnableEvents = False

With FileDialog
  .Filters.Clear
  .Title = "Select a Remit XML File"
  .Filters.Add "Remit XML File", "*.xml", 1
  .AllowMultiSelect = False
  If .Show = True Then
    xmlFile = .SelectedItems(1)
    xmlDoc.async = False: xmlDoc.validateOnParse = False
    xmlDoc.Load (xmlFile)
  End If
End With

If (xmlDoc.parseError.ErrorCode <> 0) Then
    MsgBox ("You have error " & xmlDoc.parseError.reason), vbCritical, Title
    Exit Sub
  ElseIf xmlDoc.DocumentElement.BaseName <> "REMITTable1" Then
    MsgBox "Грешен формат - REMITTable1", vbCritical, Title
    Exit Sub
End If

If Dir(Path, vbDirectory) = vbNullString Then
  VBA.FileSystem.MkDir (Path)
End If

newFile = Path & SelectedFile.GetBaseName(xmlFile) & "." & SelectedFile.GetExtensionName(xmlFile)
If Dir(newFile) = vbNullString Then
  FileCopy xmlFile, newFile
Else
  MsgBox "Не се позволява дублиране на данни!", vbCritical, Title
  Exit Sub
End If

s = 0
s1 = 0
c = 0
'----------------------------------- OrderList -------------------------------------
'For Each Segment In xmlDoc.SelectNodes("//OrderList/OrderReport")
  For Each Segment1 In xmlDoc.SelectNodes("//OrderList/OrderReport/priceIntervalQuantityDetails")

    i = LastRow

    ws.Range("A" & i).Value = "OrderReport"
    ws.Range("AQ" & i).Value = newFile

    If reportingEntityID(1, 2) <> "" Then
      ws.Range(reportingEntityID(1, 2) & i).Value = node_check(xmlDoc.SelectSingleNode(reportingEntityID(1, 1)))
    End If
  
    For Item = 1 To OrderList.Rows.Count
      If OrderList(Item, 2) <> "" And node_check(Segment1.ParentNode.SelectSingleNode(OrderList(Item, 1))) <> "" Then
        ws.Range(OrderList(Item, 2) & i).Value = node_check(Segment1.ParentNode.SelectSingleNode(OrderList(Item, 1)))
        s = s + 1
      End If
    Next Item
  
    For Item1 = 1 To QuantityDetails.Rows.Count
      If QuantityDetails(Item1, 2) <> "" And node_check(Segment1.SelectSingleNode(QuantityDetails(Item1, 1))) <> "" Then
        ws.Range(QuantityDetails(Item1, 2) & i).Value = node_check(Segment1.SelectSingleNode(QuantityDetails(Item1, 1)))
        s1 = s1 + 1
      End If
    Next Item1

    contractId = node_check(Segment1.ParentNode.SelectSingleNode("contractInfo/contractId"))
Debug.Print contractId
    For Each Key In xmlDoc.SelectNodes("//contractList/contract/contractId[.='" & contractId & "']")
      For Item2 = 1 To contractList.Rows.Count
        If contractList(Item2, 2) <> "" And node_check(Key.ParentNode.SelectSingleNode(contractList(Item2, 1))) <> "" Then
          ws.Range(contractList(Item2, 2) & i).Value = node_check(Key.ParentNode.SelectSingleNode(contractList(Item2, 1)))
          c = c + 1
        End If
      Next Item2
    Next Key
  Next Segment1
'Next Segment
Debug.Print "segment: " & s, "segment1: " & s1, "contract:" & c

'---------------------------------- TradeReport ------------------------------------
For Each Segment In xmlDoc.SelectNodes("//TradeList/TradeReport")

  i = LastRow

  ws.Range("A" & i).Value = "TradeReport"
  ws.Range("AQ" & i).Value = newFile

  If reportingEntityID(1, 2) <> "" Then
    ws.Range(reportingEntityID(1, 2) & i).Value = node_check(xmlDoc.SelectSingleNode(reportingEntityID(1, 1)))
  End If
  
  For Item = 1 To TradeList.Rows.Count
    If TradeList(Item, 2) <> "" Then
      ws.Range(TradeList(Item, 2) & i).Value = node_check(Segment.SelectSingleNode(TradeList(Item, 1)))
    End If
  Next Item

  contractId = node_check(Segment.SelectSingleNode("contractInfo/contractId"))

  For Each Key In xmlDoc.SelectNodes("//contractList/contract/contractId[.='" & contractId & "']")
    For Item1 = 1 To contractList.Rows.Count
      If contractList(Item1, 2) <> "" Then
        ws.Range(contractList(Item1, 2) & i).Value = node_check(Key.ParentNode.SelectSingleNode(contractList(Item1, 1)))
      End If
    Next Item1
  Next Key
Next Segment

If LastRow > 30 Then
  Application.Goto ws.Cells(LastRow - 30, "A"), True
End If

ws.Cells(LastRow, "A").Select

Application.ScreenUpdating = True
Application.EnableEvents = True

'MsgBox "Импортиранетo завърши успешно!", vbQuestion, Title
Debug.Print Timer - ST
End Sub
