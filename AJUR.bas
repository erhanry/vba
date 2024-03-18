Attribute VB_Name = "AJUR"
' Author Erhan Ysuf

Public Const Pass As String = ""

Public Property Get ws() As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
End Property

Public Property Get Path() As String
    Path = ThisWorkbook.Path & "\Export\"
End Property

Sub CSV(Name As String, Zona As String)

Dim Msg As Boolean: Msg = False
Dim Order As Range: Set Order = ws.Range(Zona)
Dim File As String: File = Name & Format(Date, "-dd.mm.yyyy") & ".csv"

    With Application
      .ScreenUpdating = False
      .DisplayAlerts = False
      .EnableEvents = False
    End With
    ws.Unprotect Password:=Pass
    
    If Dir(Path, vbDirectory) = vbNullString Then
       VBA.FileSystem.MkDir (Path)
    End If

    With Order
      If WorksheetFunction.Sum(.Columns.Item(4)) > 0 And WorksheetFunction.Sum(.Columns.Item(5)) > 0 Then
        Open Path & File For Output As #1
          For i = 1 To .Rows.Count
            If Not IsEmpty(.Cells(i, 4)) And .Cells(i, 4) <> 0 And Not IsEmpty(.Cells(i, 5)) And .Cells(i, 5) <> 0 Then
            Print #1, Join(Application.Transpose(Application.Transpose(.Rows.Item(i).Value)), ";")
            End If
          Next
        Close #1
        Msg = True
      End If
    End With

    ws.Protect Password:=Pass, DrawingObjects:=False
    With Application
      .ScreenUpdating = True
      .DisplayAlerts = True
      .EnableEvents = True
    End With

    If Msg Then
      MsgBox "Åêñïîðòèðàíåòî íà äàííè çàâúðøè óñïåøíî!" & vbNewLine & vbCrLf & File, vbQuestion, "Export Ajur CSV File."
    Else
      MsgBox "Ïðàçíà Òàáëèöà!", vbCritical, "Export Ajur CSV File."
    End If

End Sub


Sub XLS(Name As String, Zona As String, Service As Variant)

Dim Msg As Boolean: Msg = False
Dim Order As Range: Set Order = ws.Range(Zona)
Dim File_name As String
Dim File As String
Dim File_export As Workbook
Dim Data_arr(1 To 25, 1 To 3) As Variant

    With Application
      .ScreenUpdating = False
      .DisplayAlerts = False
      .EnableEvents = False
    End With
    ws.Unprotect Password:=Pass
    
    If Dir(Path, vbDirectory) = vbNullString Then
       VBA.FileSystem.MkDir (Path)
    End If
    
    With Order
    For Each Service_Key In Service

    File_name = Name & "_usl_" & Service_Key & Format(Date, "-dd.mm.yyyy")
    On Error Resume Next
        Workbooks(File_name).Close SaveChanges:=False
        Kill Path & File_name & ".xls"
    On Error GoTo 0

    If WorksheetFunction.SumIf(.Columns.Item(1), Service_Key, .Columns.Item(6)) > 0 Then
        j = 1
        Erase Data_arr
        For i = 1 To .Rows.Count
            If .Cells(i, 1) = Service_Key And .Cells(i, 6) > 0 Then
                Data_arr(j, 1) = .Cells(i, 2).Value
                Data_arr(j, 2) = .Cells(i, 5).Value
                Data_arr(j, 3) = .Cells(i, 6).Value
                 j = j + 1
            End If
        Next i
        Workbooks.Add.SaveAs Filename:=Path & File_name & ".xls", _
                            ReadOnlyRecommended:=False, _
                            FileFormat:=56, _
                            CreateBackup:=False, _
                            AccessMode:=xlExclusive, _
                            ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        File = File & vbNewLine & vbCrLf & File_name & ".xls"
        Set File_export = ActiveWorkbook
        File_export.Worksheets(1).Range("A1").Resize(UBound(Data_arr, 1), UBound(Data_arr, 2)) = Data_arr()
        File_export.Close SaveChanges:=True
        Msg = True
    End If
    Next Service_Key
    End With

    ws.Protect Password:=Pass, DrawingObjects:=False
    With Application
      .ScreenUpdating = True
      .DisplayAlerts = True
      .EnableEvents = True
    End With
    If Msg Then
        MsgBox "Åêñïîðòèðàíåòî íà äàííè çàâúðøè óñïåøíî!" & File, vbQuestion, "Export Ajur XLS File."
    Else
        MsgBox "Ïðàçíà Òàáëèöà!", vbCritical, "Export Ajur XLS File."
    End If
End Sub
