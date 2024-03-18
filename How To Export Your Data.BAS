Attribute VB_Name = "Module1"
Option Explicit

Sub ExportData()
'http://www.howtoexcel.org/
'John MacDougall 2017-05-07

'Declare variables
Dim ArrayItem As Long
Dim ws As Worksheet
Dim ArrayOfUniqueValues As Variant
Dim SavePath As String
Dim ColumnHeadingInt As Long
Dim ColumnHeadingStr As String
Dim rng As Range

'Set the worksheet to
Set ws = Sheets("Data")

'Set the save path for the files created
SavePath = Range("FolderPath")

'Set variables for the column we want to separate data based on
ColumnHeadingInt = WorksheetFunction.Match(Range("ExportCriteria").Value, Range("Data[#Headers]"), 0)
ColumnHeadingStr = "Data[[#All],[" & Range("ExportCriteria").Value & "]]"

'Turn off screen updating to save runtime
Application.ScreenUpdating = False

'Create a temporary list of unique values from the column we want to
'separate our data based on
Range(ColumnHeadingStr & "").AdvancedFilter Action:=xlFilterCopy, _
    CopyToRange:=Range("UniqueValues"), Unique:=True

'Sort our temporary list of unique values
ws.Range("UniqueValues").EntireColumn.Sort Key1:=ws.Range("UniqueValues").Offset(1, 0), _
    Order1:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:=False, _
    Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

'Add unique field values into an array
'ArrayOfUniqueValues = Application.WorksheetFunction.Transpose(ws.Range("IV2:IV" & Rows.Count).SpecialCells(xlCellTypeConstants))
ArrayOfUniqueValues = Application.WorksheetFunction.Transpose(ws.Range("UniqueValues").EntireColumn.SpecialCells(xlCellTypeConstants))

'Delete the temporary values
ws.Range("UniqueValues").EntireColumn.Clear

'Loop through our array of unique field values, copy paste into new workbooks and save
For ArrayItem = 1 To UBound(ArrayOfUniqueValues)
    ws.ListObjects("Data").Range.AutoFilter Field:=ColumnHeadingInt, Criteria1:=ArrayOfUniqueValues(ArrayItem)
    ws.Range("Data[#All]").SpecialCells(xlCellTypeVisible).Copy
    Workbooks.Add
    Range("A1").PasteSpecial xlPasteAll
    ActiveWorkbook.SaveAs SavePath & ArrayOfUniqueValues(ArrayItem) & Format(Now(), " YYYY-MM-DD hhmmss") & ".xlsx", 51
    ActiveWorkbook.Close False
    ws.ListObjects("Data").Range.AutoFilter Field:=ColumnHeadingInt
Next ArrayItem

ws.AutoFilterMode = False
MsgBox "Finished exporting!"
Application.ScreenUpdating = True
    
End Sub





