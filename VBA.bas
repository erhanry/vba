----------------------------------------
Sub Multi_Copy_Sheet()
Dim n As Integer
Dim i As Integer
On Error Resume Next
    n = InputBox("How many copies do you want to make?")
    If n > 0 Then
        For i = 1 To n
           Sheets("Main").Copy After:=ActiveWorkbook.Sheets(Worksheets.Count)
           ActiveSheet.Name = Format(i, "00")
        Next i
    End If
End Sub
----------------------------------------
Sub Hide_Sheet()
Dim i As Integer
	For i = 2 To Worksheets.Count
		Sheets(i).Visible = False
	Next i
End Sub
----------------------------------------
Sub Visible_Sheet()
	Dim i As Integer
	For i = 1 To Worksheets.Count
		Sheets(i).Visible = True
	Next i
End Sub
----------------------------------------
Sub Copy_Select(Zona As String)
	ActiveSheet.Select
	Range(Zona).Select
	Selection.Copy
End Sub
----------------------------------------
Sub Paste_Values(Zona As String)
   Range(Zona).Select
   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
----------------------------------------
Sub cop()
Dim i As Integer
	For i = 1 To 31
		Sheets(Format(i, "00")).Unprotect Password:="113830"
		Sheets(Format(i, "00")).Range("C33").Formula = "=Schedules!E752"
		Sheets(Format(i, "00")).Protect Password:="113830", AllowFormattingCells:=True, DrawingObjects:=False
	Next i
End Sub
----------------------------------------
=INDIRECT("Duty[Дежурни]")


Sub cop()
Dim i As Integer
    For i = 2 To 31
        Sheets(Format(i, "00")).Unprotect Password:="113830"
        Sheets(Format(i, "00")).Columns("ET:EZ").EntireColumn.Delete
        Sheets(Format(i, "00")).Protect Password:="113830", AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub

Sub cop()
Dim i As Integer
    For i = 2 To 31
        Sheets(Format(i, "00")).Unprotect Password:="113830"
        Sheets(Format(i, "00")).Range("B4").Interior.Color = RGB(189, 215, 238)
        Sheets(Format(i, "00")).Range("B4").Locked = True
        Sheets(Format(i, "00")).Range("B4") = "='" + Format(i - 1, "00") + "'!B28"
        Sheets(Format(i, "00")).Protect Password:="113830", AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub

Sub cop()
Dim i As Integer
    For i = 2 To 31
        Sheets(Format(i, "00")).Unprotect Password:="113830"
        Sheets(Format(1, "00")).Range("ET:EW").Copy Sheets(Format(i, "00")).Range("ET:EW")
        Sheets(Format(i, "00")).Protect Password:="113830", AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub

Sub opcom()

ActiveSheet.Protect Password:="113830", AllowFormattingCells:=True, DrawingObjects:=False
Application.ScreenUpdating = True

Dim this_day As Long: this_day = CDbl(ThisWorkbook.Sheets("Config").Range("start_date")) + CDbl(ThisWorkbook.ActiveSheet.Name) - 1
Dim Price_Array(0 To 24) As Variant
Dim i As Integer: i = 0

Url = "https://www.opcom.ro/opcom/rapoarte/pzu/export_xml_PIPsiVolTran.php?zi=" & _
    Day(this_day) & "&luna=" & Month(this_day) & "&an=" & Year(this_day) & "&limba=en"

Set xmldoc = New MSXML2.DOMDocument60
xmldoc.async = False
xmldoc.Load (Url)

For Each List In xmldoc.SelectNodes("//resultset/Detail")
  Price_Array(i) = Round(List.SelectSingleNode("Price").Text * 1.95583, 2)
  i = i + 1
Next List

ThisWorkbook.ActiveSheet.Range("K1").Value = "Opcom Price"
ThisWorkbook.ActiveSheet.Range("K5:K28").Value = Application.Transpose(Price_Array)

Set xmldoc = Nothing
Set List = Nothing

    ActiveSheet.Protect Password:="113830", AllowFormattingCells:=True, DrawingObjects:=False
    Application.ScreenUpdating = True

End Sub


Sub orderImport()
  Open ThisWorkbook.Path & "\template.csv" For Output As #1
  
  Print #1, Join(Array("Area", "Portfolio", "Product", "Direction", "Quantity", "Price", "Type", "Label"), ";")

  For i = 1 To 2
    Print #1, Join(Array("BG", "ENERGO-PRO BULGARIA EAD", "PH-20210521-22", "Sell", "0.1", "100", "Limit", "Import"), ";")
  Next i

  Close #1
End Sub

Sheets(Format(i, "00")).Shapes("Group 10").NAME = "Group 1"


Sub Shapes_ALL()
  Debug.Print ActiveSheet.Shapes.Count
  For x = 1 To ActiveSheet.Shapes.Count
    Debug.Print ActiveSheet.Shapes(x).Name
    Debug.Print ActiveSheet.Shapes(x).Left
    Debug.Print ActiveSheet.Shapes(x).Top
    Debug.Print ActiveSheet.Shapes(x).Width
    Debug.Print ActiveSheet.Shapes(x).Height
  Next
End Sub

Sheets(Format(i, "00")).Range("C33").Formula = "=Schedules!E752"
Sheets(Format(i, "00")).Range("C33").HorizontalAlignment = xlCenter
Sheets(Format(i, "00")).Range("C33").NumberFormat = "#,##0"



Sub Shapes_ALL()
i = 5
j = 7
k = 30
Debug.Print ActiveSheet.Shapes.Count
  For x = 1 To ActiveSheet.Shapes.Count
   ActiveSheet.Shapes(x).OnAction = "'MMS ""H" & i & """, ""C" & j & ":M" & k & """'"
   Debug.Print ActiveSheet.Shapes(x).OnAction
i = i + 32
j = j + 32
k = k + 32
Next
End Sub
Sub cop()
Dim i As Integer
    For i = 1 To 31
        Sheets(Format(i, "00")).Unprotect Password:="113830"
        Sheets(Format(i, "00")).Range("L35:L36").Copy
        Sheets(Format(i, "00")).Range("BX36").PasteSpecial
        Sheets(Format(i, "00")).Range("CA36").PasteSpecial
        Sheets(Format(i, "00")).Range("CD36").PasteSpecial
        Sheets(Format(i, "00")).Range("CG36").PasteSpecial
        Sheets(Format(i, "00")).Range("CJ36").PasteSpecial
        Sheets(Format(i, "00")).Range("CM36").PasteSpecial
        Sheets(Format(i, "00")).Range("CP36").PasteSpecial
        Sheets(Format(i, "00")).Range("CS36").PasteSpecial
        Sheets(Format(i, "00")).Range("CV36").PasteSpecial
        Sheets(Format(i, "00")).Range("CY36").PasteSpecial
        Sheets(Format(i, "00")).Range("DB36").PasteSpecial
        Sheets(Format(i, "00")).Range("DE36").PasteSpecial
        Sheets(Format(i, "00")).Range("DH36").PasteSpecial
        Sheets(Format(i, "00")).Range("DK36").PasteSpecial
        Sheets(Format(i, "00")).Range("DN36").PasteSpecial
        Sheets(Format(i, "00")).Range("DQ36").PasteSpecial
        Sheets(Format(i, "00")).Range("DT36").PasteSpecial
        Sheets(Format(i, "00")).Range("DW36").PasteSpecial
        Sheets(Format(i, "00")).Range("DZ36").PasteSpecial
        Sheets(Format(i, "00")).Range("EC36").PasteSpecial
        Sheets(Format(i, "00")).Range("EF36").PasteSpecial
        Sheets(Format(i, "00")).Columns("CB:CJ").EntireColumn.Hidden = True
        Sheets(Format(i, "00")).Range("A1:A3").Select
        Sheets(Format(i, "00")).Protect Password:="113830", AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub

Sub cop()
Dim i As Integer
    For i = 1 To 31
        Sheets(Format(i, "00")).Unprotect Password:="113830"
        Sheets( "01").Range("CK31").Copy
        Sheets(Format(i, "00")).Range("CK31").PasteSpecial
        Sheets( "01").Range("CM31").Copy
        Sheets(Format(i, "00")).Range("CM31").PasteSpecial
        Sheets("01").Range("A1:A3").Copy
        Sheets(Format(i, "00")).Range("A1:A3").PasteSpecial Paste:=xlPasteValues
        Sheets(Format(i, "00")).Protect Password:="113830", AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub


Sub cop()
Dim i As Integer
    For i = 1 To 31
        Sheets(Format(i, "00")).Unprotect Password:=Pass
        Sheets(Format(i, "00")).Range("GF4").Value = 8000
        Sheets(Format(i, "00")).Protect Password:=Pass, AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub

Sub cop()
Dim i As Integer
    For i = 1 To 31
        Sheets(Format(i, "00")).Unprotect Password:=Pass
        Sheets("01").Range("A1:A3").Copy
        Sheets(Format(i, "00")).Range("A1:A3").PasteSpecial Paste:=xlPasteValues
        Sheets(Format(i, "00")).Protect Password:=Pass, AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub

Sub Macro1()
 For i = 720 To 1 Step -1
    Cells(i + 1, 1).EntireRow.Insert
     Cells(i + 1, 1).Value = Cells(i, 1).Value
    Cells(i + 1, 1).EntireRow.Insert
     Cells(i + 1, 1).Value = Cells(i, 1).Value
    Cells(i + 1, 1).EntireRow.Insert
     Cells(i + 1, 1).Value = Cells(i, 1).Value
    Next i
End Sub

Sub cop1()
Dim i As Integer
    For i = 1 To 1
        Sheets(Format(i, "00")).Unprotect Password:=Pass
        Sheets(Format(i, "00")).Range("U:W").EntireColumn.Hidden = True
        Sheets(Format(i, "00")).Range("AV:AX").EntireColumn.Hidden = True
        Sheets(Format(i, "00")).Range("BV:EJ").EntireColumn.Hidden = True
        Sheets(Format(i, "00")).Range("EZ1").EntireColumn.Copy Sheets(Format(i, "00")).Range("EP1").EntireColumn
        Sheets(Format(i, "00")).Range("EP1:EP27").Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        Sheets(Format(i, "00")).Protect Password:=Pass, AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub

Sub cop1()
Dim i As Integer
    For i = 1 To 31
        Sheets(Format(i, "00")).Unprotect Password:=Pass
		Sheets(Format(i, "00")).Range("BV4,BY4,CK4,CN4,CQ4,CT4,CW4,CZ4,DC4,DF4,DR4,EA4").Value = 0
        Sheets(Format(i, "00")).Range("BV:EJ").EntireColumn.Hidden = True
        Sheets(Format(i, "00")).Range("EY1").EntireColumn.Copy Sheets(Format(i, "00")).Range("EP1").EntireColumn
        Sheets(Format(i, "00")).Range("EP1:EP27").Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        Sheets(Format(i, "00")).Range("BU1:BU27").Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        Sheets(Format(i, "00")).Range("EP1:EP27").Borders(xlEdgeRight).Weight = xlMedium
        Sheets(Format(i, "00")).Range("BU1:BU27").Borders(xlEdgeRight).Weight = xlMedium
        Sheets(Format(i, "00")).Range("EZ16:EZ39").UnMerge
        Sheets(Format(i, "00")).Range("EZ16:EZ27").Merge
		Sheets(Format(i, "00")).Range("EZ28:FZ39").Delete Shift:=xlUp
		Sheets(Format(i, "00")).Range("EZ28").Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
		'Sheets(Format(i, "00")).Range("a1").Select
        Sheets(Format(i, "00")).Protect Password:=Pass, AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub

Sub soho()
Dim k As Integer
k = 3
For i = 1 + 3 To 24 + 3
    For j = 1 To 4
        ActiveSheet.Cells(12, k).FormulaR1C1 = "=EPB!R22C" & i
        k = k + 1
    Next j
Next i
End Sub

Sub soho()
Dim k As Integer
k = 3
For i = 1 + 3 To 25 + 3
    For j = 1 To 4
        ActiveSheet.Cells(12, k).FormulaR1C1 = "=ROUND(EPB!R21C" & i & "/4,3)"
        k = k + 1
    Next j
Next i
End Sub


Sub mat2()
Dim i As Integer
    For i = 1 To 1
        Sheets(Format(i, "00")).Unprotect Password:=Pass
        Sheets(Format(i, "00")).Range("AT4:AT28").Locked = True
        Sheets(Format(i, "00")).Protect Password:=Pass, AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub

Sub soho()
Dim K As Integer
K = 3
For i = 1 To 25
K = K + 1
    'ActiveSheet.Cells(22, i + 157).Formula = "=$Z$" & K & "+$AA$" & K & "+$AB$" & K & "+$AM$" & K & "+$AN$" & K & "+$AO$" & K & "+$AP$" & K & "-$AC$" & K & "-$AQ$" & K
    ActiveSheet.Cells(23, i + 157).Formula = "=$AT$" & K & "+$AU$" & K & "+$AV$" & K & "+$AW$" & K & "+$BA$" & K & "+$BB$" & K & "+$BC$" & K & "+$BD$" & K & "+$BH$" & K & "+$BI$" & K & "+$BJ$" & K & "+$BK$" & K & "-$AX$" & K & "-$BE$" & K & "-$BL$" & K & "+$BO$" & K & "+$BP$" & K & "+$BQ$" & K & "+$BR$" & K & "-$BS$" & K
Next i
End Sub


Sub cop()
Dim i As Integer
    For i = 2 To 31
        Sheets(Format(i, "00")).Unprotect Password:=Pass
        Sheets(Format(1, "00")).Range("EZ:FZ").Copy Sheets(Format(i, "00")).Range("ET:EW")
        Sheets(Format(i, "00")).Protect Password:=Pass, AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub


Sub gfmes()
Dim T, L As Double

T = Sheets(Format(1, "00")).Shapes(2).Top
L = Sheets(Format(1, "00")).Shapes(2).Left

For i = 2 To 31
    Sheets(Format(i, "00")).Unprotect Password:=Pass
    Sheets(Format(1, "00")).Shapes(2).Copy
    Sheets(Format(i, "00")).Shapes(2).Delete
    Sheets(Format(i, "00")).Paste
    Sheets(Format(i, "00")).Shapes(2).Top = T
    Sheets(Format(i, "00")).Shapes(2).Left = L
    Sheets(Format(i, "00")).Range("A1").Copy
    Sheets(Format(i, "00")).Range("A1").PasteSpecial Paste:=xlPasteValues
Sheets(Format(i, "00")).Protect Password:=Pass, AllowFormattingCells:=True, DrawingObjects:=False
 Next i
End Sub


Sub cop()
Dim i As Integer
    For i = 1 To 31
        Sheets(Format(i, "00")).Unprotect Password:=Pass
        Sheets(Format(i, "00")).Range("C35").Formula = "=den()"
        Sheets(Format(i, "00")).Protect Password:=Pass, AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub

Sub PasswordBreaker()
  'Breaks worksheet password protection.

  Dim i As Integer, j As Integer, k As Integer
  Dim l As Integer, m As Integer, n As Integer
  Dim i1 As Integer, i2 As Integer, i3 As Integer
  Dim i4 As Integer, i5 As Integer, i6 As Integer

  On Error Resume Next

  For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
  For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
  For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
  For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126

 ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)

    If ActiveSheet.ProtectContents = False Then
        Debug.Print "Password is " & Chr(i) & Chr(j) & _
          Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
          Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)

        Exit Sub

    End If

  Next: Next: Next: Next: Next: Next
  Next: Next: Next: Next: Next: Next
End Sub


Sub cont()
Dim i, j, k As Integer
k = 2
    For i = 3 To 747
        For j = 1 To 4
        k = k + 1
        ThisWorkbook.Worksheets("Sheet2").Cells(k, 3).Formula = "=Sheet1!C" & i & "/4"
        ThisWorkbook.Worksheets("Sheet2").Cells(k, 4).Formula = "=Sheet1!D" & i & "/4"
        ThisWorkbook.Worksheets("Sheet2").Cells(k, 5).Formula = "=Sheet1!E" & i & "/4"
        ThisWorkbook.Worksheets("Sheet2").Cells(k, 6).Formula = "=Sheet1!F" & i & "/4"
        ThisWorkbook.Worksheets("Sheet2").Cells(k, 7).Formula = "=Sheet1!G" & i & "/4"
        ThisWorkbook.Worksheets("Sheet2").Cells(k, 8).Formula = "=Sheet1!H" & i & "/4"
        ThisWorkbook.Worksheets("Sheet2").Cells(k, 9).Formula = "=Sheet1!I" & i & "/4"
        ThisWorkbook.Worksheets("Sheet2").Cells(k, 10).Formula = "=Sheet1!J" & i & "/4"
        ThisWorkbook.Worksheets("Sheet2").Cells(k, 11).Formula = "=Sheet1!K" & i & "/4"
        Next j
    Next i
End Sub

Sub FitComments()
 
Dim xComment As Comment
For Each xComment In Application.ActiveSheet.Comments
    xComment.Shape.TextFrame.AutoSize = True
Next
End Sub


Sub cop()
Dim i As Integer
    For i = 2 To 31
        Sheets(Format(i, "00")).Unprotect Password:=Pass
        Sheets(Format(1, "00")).Range("F35:I36").Copy Sheets(Format(i, "00")).Range("F35:I36")
        For j = 4 To 27
            Sheets(Format(i, "00")).Range("CI" & j).Formula = "=AG" & j & "+BA" & j & "+BO" & j
        Next j
        Sheets(Format(i, "00")).Protect Password:=Pass, AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub
 
Sub eur()
For i = 1 To 31
    With Sheets(Format(i, "00"))
        .Unprotect Password:=Pass
        
        .Range("CY4").Value = -500
        .Range("DG4").Value = -500
        .Range("DA4").Value = 25
        .Range("DI4").Value = 24
        .Range("DC4").Value = 125
        .Range("DD4").Value = 4000
        .Range("DJ4").Value = 4000
        .Range("K4:L28,F31,H31,K30:L31,K33:L34").NumberFormat = "# ##0.00\ [$И-x-euro1]"
        .Range("F35:I36").ClearContents
        .Range("M31").Formula = "=N33*N32*24"
        .Range("M31").NumberFormat = """ѕрогн: ""#,##0.0"" MWh"""
        .Range("F31").Value = "=IFERROR(F30/E30,""" & ChrW(&H58D) & """)"
        .Range("H31").Value = "=IFERROR(H30/G30,""" & ChrW(&H58D) & """)"
        .ChartObjects("Chart 1").Chart.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "# ##0\ [$И-x-euro1]"
        .Range("N4:O28,U4:U28,AA4:AB28,AH4:AI28,AO4:AP28,AV4:AV28,BB4:BC28,BI4:BJ28,BP4:BQ28").Locked = True
        
        .Protect Password:=Pass, AllowFormattingCells:=True, DrawingObjects:=False
    End With
Next i
End Sub