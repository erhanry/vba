Attribute VB_Name = "Intraday"
Sub Creare()

Dim Order As Range
Dim msg As Boolean: msg = False

Application.ScreenUpdating = False
Application.EnableEvents = False

With ThisWorkbook.ActiveSheet
  .Unprotect Password:="113830"
  Set Order = .Range("B1:I25")

  If .Range("F27") > 0 And .Range("H27") > 0 Then
    Open ThisWorkbook.Path & "\Orders.csv" For Output As #1
      For i = Order.Row To Order.Rows.Count
        If Not IsEmpty(Order(i - Order.Row + 1, 5)) And Not IsEmpty(Order(i - Order.Row + 1, 6)) Then
          Print #1, Join(Application.Transpose(Application.Transpose(.Range("B" & i & ":I" & i).Value)), ";")
        End If
      Next i
    Close #1
    msg = True
  End If

  .Protect Password:="113830"
End With

Application.ScreenUpdating = True
Application.EnableEvents = True

If msg Then
  MsgBox "The Order is successfully created!" & vbNewLine & vbCrLf & ThisWorkbook.Path & "\Orders.csv", vbQuestion, "Intraday order generator."
Else
  MsgBox "The form is blank!", vbCritical, "Intraday order generator."
End If

End Sub

Sub Clear()
  ThisWorkbook.ActiveSheet.Range("F2:G25").ClearContents
End Sub
