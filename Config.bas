Attribute VB_Name = "Config"
' Campany Energo-Pro Bulgaria
' Author Erhan Ysuf

Function den()
    Application.Volatile
    den = StrConv(Format(CInt(Application.Caller.Worksheet.Name) + CDbl(ThisWorkbook.Sheets("Config").Range("start_date")) - 1, "dddd"), vbProperCase)
End Function
