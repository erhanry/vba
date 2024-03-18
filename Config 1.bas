Sub mat()
Dim i As Integer
    For i = 1 To 31
        Sheets(Format(i, "00")).Unprotect Password:=Pass
        Sheets(Format(i, "00")).Range("U5:U28").Locked = False
        Sheets(Format(i, "00")).Range("AV5:AV28").Locked = False
        Sheets(Format(i, "00")).Range("GB3:GF32").Copy Sheets(Format(i, "00")).Range("GH3")
        Sheets(Format(i, "00")).Range("GB3").Value = "IBEX Auctions - Примерна матрица за ЕПБ"
        Sheets(Format(i, "00")).Range("GH3").Value = "IBEX Auctions - Примерна матрица за ПБЕ"
        Sheets(Format(i, "00")).Range("GE5:GE28").FormulaR1C1 = "=-ROUND(RC[-172]+RC[-160]+RC[-153]+RC[-146]+RC[-132]+RC[-125]+RC[-118],1)"
        Sheets(Format(i, "00")).Range("GF5:GF28").FormulaR1C1 = "=-ROUND(RC[-173]+RC[-161]+RC[-154]+RC[-147]+RC[-133]+RC[-126]+RC[-119],1)"
        Sheets(Format(i, "00")).Range("GK5:GK28").FormulaR1C1 = "=-ROUND(RC[-172]+RC[-145],1)"
        Sheets(Format(i, "00")).Range("GL5:GL28").FormulaR1C1 = "=-ROUND(RC[-173]+RC[-146],1)"
        Sheets(Format(i, "00")).Range("GK30").Formula = "=SUM($GK$5:$GK$28)"
        Sheets(Format(i, "00")).Range("GK31").Formula = "=MIN($GK$5:$GK$28)"
        Sheets(Format(i, "00")).Range("GK32").Formula = "=MAX($GK$5:$GK$28)"
        Sheets(Format(i, "00")).Range("GK4").Value = 48
        Sheets(Format(i, "00")).Protect Password:=Pass, AllowFormattingCells:=True, DrawingObjects:=False
    Next i
End Sub


Sub d()
Debug.Print Sheets(Format(4, "00")).Shapes(2).GroupItems(27).TextFrame.Characters.Text
Debug.Print Sheets(Format(4, "00")).Shapes(2).GroupItems(27).OnAction
Debug.Print Sheets(Format(4, "00")).Shapes(2).GroupItems(27).Top
Debug.Print Sheets(Format(4, "00")).Shapes(2).GroupItems(27).Left

End Sub
