Attribute VB_Name = "Remit"
Private Sub Button1_Click()
  Dim fd As Office.FileDialog
  Dim xDoc As Object

  Set fd = Application.FileDialog(msoFileDialogFilePicker)
  With fd
    .Filters.Clear
    .Title = "Избери Remit XML файл"
    .Filters.Add "Remit XML File", "*.xml", 1
    .AllowMultiSelect = False

    If .Show = True Then
      xmlFileName = .SelectedItems(1)
      ' Process XML File
      Set xDoc = CreateObject("MSXML2.DOMDocument")
      xDoc.async = False: xDoc.validateOnParse = False
      xDoc.Load (xmlFileName)
      
      
      ' Get Root Node
      For Each TradeList In xDoc.SelectNodes("//OrderList/OrderReport")
      i = LastRow
      j = 0
        For Each TradeReport In TradeList.ChildNodes
             Cells(i, "A").Offset(, j).Value = TradeReport.Text
             j = j + 1
        Next TradeReport
      Next TradeList


      ' Get Root Node
      For Each TradeList In xDoc.SelectNodes("//TradeList/TradeReport")
      i = LastRow
      j = 0
        For Each TradeReport In TradeList.ChildNodes
             Cells(i, "A").Offset(, j).Value = TradeReport.Text
             j = j + 1
        Next TradeReport
      Next TradeList




     End If
  End With
  
End Sub


Function LastRow()
    LastRow = Cells(1, "A").End(xlDown).Row + 1
End Function

