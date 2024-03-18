Attribute VB_Name = "Devision"
Sub Schedule_KBG()
ActiveSheet.Unprotect Password:="113830"
Call Schedule_Copy(Sheets("Config").Range("Devision_Create_Dir").Value, _
                    Sheets("Config").ListObjects("Devision_Create").DataBodyRange)
ActiveSheet.Protect Password:="113830", AllowFormattingCells:=True, DrawingObjects:=False
End Sub

Function Schedule_Copy(Set_Path As String, MyArray As Variant)

Dim answer As Integer
Const title As String = "Генериране на графици."

FolderPath = ThisWorkbook.Path & Set_Path
Application.ScreenUpdating = False
Application.EnableEvents = False

If ActiveSheet.Range(ThisWorkbook.Sheets("Config").Range("balans")).Value <> 0 Then
answer = MsgBox("Имате небаланс, искатели да продължите?", vbQuestion + vbYesNo)
 End If

If answer = vbNo Then Exit Function

If Dir(FolderPath, vbDirectory) = vbNullString Then
    MsgBox "Директорията не е намерена!", vbCritical, title
    Exit Function
End If

For Item = 1 To MyArray.Rows.Count
If Dir(FolderPath & MyArray(Item, 1) & ".xlsx", vbDirectory) = vbNullString Then
    MsgBox "Файла: """ & MyArray(Item, 1) & """ не е намерен!" & vbNewLine & vbCrLf & "Генерирането не е започнат!", vbCritical, title
    Exit Function
End If
Next Item

For Item = 1 To MyArray.Rows.Count
   Range(MyArray(Item, 2)).Copy
   Workbooks.Open FolderPath & MyArray(Item, 1) & ".xlsx", Format:=3, Origin:=3
   Range("B4").Value = CDbl(ThisWorkbook.Sheets("Config").Range("start_date")) - 1 + CInt(ThisWorkbook.ActiveSheet.Name)
   Range("B6").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   Range("B4").NumberFormat = "d.m.yyyy"
   ActiveWorkbook.Close SaveChanges:=True
Next Item

AutoFilterMode = False
Application.CutCopyMode = False
Application.ScreenUpdating = True
Application.EnableEvents = True
MsgBox "Успешно са генерирани " & Item - 1 & " файла.", vbQuestion, title
End Function
