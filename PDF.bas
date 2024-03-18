Attribute VB_Name = "MMS"
' Campany Energo-Pro Bulgaria
' Author Erhan Ysuf

Sub TPS_DAM_EPB()
    Call MMS("O4:O27,X4:X27,AE4:AE27,AL4:AL27,AW4:AW27,BD4:BD27,BK4:BK27", 2)
End Sub

Sub TPS_IDM_EPB()
    Call MMS("P4:Q27,Y4:Z27,AF4:AG27,AM4:AN27,AX4:AY27,BE4:BF27,BL4:BM27", 11)
End Sub

Sub TPS_DAM_SB()
    Call MMS("N4:O27", 26)
End Sub

Sub TPS_IDM_SB()
    Call MMS("P4:Q27", 29)
End Sub

Sub PPS_SB()
    Call MMS("N4:Q27", 32)
End Sub

Sub TPS_DAM_ST()
    Call MMS("W4:X27", 38)
End Sub

Sub TPS_IDM_ST()
    Call MMS("Y4:Z27", 41)
End Sub

Sub TPS_DAM_PT()
    Call MMS("AC4:AE27", 44)
End Sub

Sub TPS_IDM_PT()
    Call MMS("AF4:AG27", 48)
End Sub

Sub PPS_PT()
    Call MMS("AC4:AG27", 51)
End Sub

Sub TPS_DAM_KP()
    Call MMS("AJ4:AL27", 58)
End Sub

Sub TPS_IDM_KP()
    Call MMS("AM4:AN27", 62)
End Sub

Sub TPS_DAM_KT()
    Call MMS("AU4:AW27", 65)
End Sub

Sub TPS_IDM_KT()
    Call MMS("AX4:AY27", 69)
End Sub

Sub TPS_DAM_SM()
    Call MMS("BB4:BD27", 72)
End Sub

Sub TPS_IDM_SM()
    Call MMS("BE4:BF27", 76)
End Sub

Sub TPS_DAM_KR()
    Call MMS("BI4:BK27", 79)
End Sub

Sub TPS_IDM_KR()
    Call MMS("BL4:BM27", 83)
End Sub


Sub MMS(Source_Range As String, col As Integer)

    Dim MMS As Worksheet
    Set MMS = Sheets("MMS")
    Const title As String = "Ãåíåðèðàíå íà MMS."
    Path = ThisWorkbook.Path & Sheets("Config").Range("MMS_Dir").Value

    If Dir(Path, vbDirectory) = vbNullString Then
        MsgBox "Äèðåêòîðèÿòà íå å íàìåðåíà!", vbCritical, title
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect Password:=""
    ThisDay = CDbl(Sheets("Config").Range("start_date")) - 1 + CInt(ActiveSheet.Name)
    MMS.Cells(1, col).Value = ThisDay
    ActiveSheet.Range(Source_Range).SpecialCells(xlCellTypeVisible).Copy
    MMS.Cells(44, col).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Folder = Format(ThisDay, "dd.mm.yyyy") & "\"
    File = MMS.Cells(2, col).Value & "_V"
    msgIdtfc = Format(ThisDay, "yyyymmdd") & "_" & Trim(MMS.Cells(3, col).Value)
    msgVersion = 1 'MMS.Cells(4, col).Value
    msgType = MMS.Cells(5, col).Value
    Gpt = MMS.Cells(6, col).Value
    'SchemeType = MMS.Cells(7, col).Value
    SenderID = Trim(MMS.Cells(8, col).Value)
    SCodingScheme = MMS.Cells(9, col).Value
    SenderRole = MMS.Cells(10, col).Value
    ReceiverID = Trim(MMS.Cells(11, col).Value)
    RCodingScheme = MMS.Cells(12, col).Value
    ReceiverRole = MMS.Cells(13, col).Value
    DatTm = Format(MMS.Cells(14, col).Value, "yyyy-mm-ddThh:mm:ssZ")
    MesTIFrom = Format(MMS.Cells(15, col).Value, "yyyy-mm-ddThh:mmZ")
    MesTITo = Format(MMS.Cells(16, col).Value, "yyyy-mm-ddThh:mmZ")

    i = col
    While Not IsEmpty(MMS.Cells(26, i).Value)
       If (MMS.Cells(23, i).Value <> "yes" Or MMS.Cells(24, i).Value > 0) And MMS.Cells(25, i).Value >= 0 Then
           check = 1
       End If
    i = i + 1
    Wend
         
    If check <> 1 Then
       Application.CutCopyMode = False
       ActiveSheet.Protect Password:="", AllowFormattingCells:=True, DrawingObjects:=False
       Application.ScreenUpdating = True
       MsgBox "Íóëåâ ãðàôèê èëè îòðèöàòåëíà ñòîéíîñò!" & vbNewLine & vbCrLf & "Ãåíåðèðàíåòî íå å çàïî÷íàò!", vbCritical, title
       Exit Sub
    End If

    If Dir(Path + Folder, vbDirectory) = vbNullString Then
       VBA.FileSystem.MkDir (Path + Folder)
    End If

    For Count = 100 To 1 Step -1
       If Dir(Path + Folder + File & Count & ".xml", vbDirectory) <> vbNullString Then
           msgVersion = Count + 1
           Exit For
       End If
    Next Count

    MMS.Cells(4, col).Value = msgVersion
    Open Path + Folder + File & msgVersion & ".xml" For Output As #1

    Print #1, "<?xml version="""; "1.0"; """"; " encoding="""; "UTF-8"""; "?>"
    Print #1, "<?xml-stylesheet type=""text/xsl"" href=""schedule-xsl.xsl""?>"
    Print #1, "<ScheduleMessage DtdVersion="""; "3"""; " DtdRelease="""; "3"""; ">"
    Print #1, "   <MessageIdentification v="""; msgIdtfc; """/>"
    Print #1, "   <MessageVersion v="""; CStr(msgVersion); """/>"
    Print #1, "   <MessageType v="""; msgType; """/>"
    Print #1, "   <ProcessType v="""; Gpt; """/>"
    Print #1, "   <ScheduleClassificationType v="""; "A01"""; "/>"
    Print #1, "   <SenderIdentification v="""; SenderID; """ codingScheme="""; SCodingScheme; """/>"
    Print #1, "   <SenderRole v="""; SenderRole; """/>"
    Print #1, "   <ReceiverIdentification v="""; ReceiverID; """ codingScheme="""; RCodingScheme; """/>"
    Print #1, "   <ReceiverRole v="""; ReceiverRole; """/>"
    Print #1, "   <MessageDateTime v="""; DatTm; """/>"
    Print #1, "   <ScheduleTimeInterval v="""; MesTIFrom; "/"; MesTITo; """/>"

    While Not IsEmpty(MMS.Cells(26, col).Value)
        If (MMS.Cells(23, col).Value <> "yes" Or MMS.Cells(24, col).Value > 0) And MMS.Cells(25, col).Value >= 0 Then
            Print #1, "   <ScheduleTimeSeries>"
            Call exportEssTimeSeries(col, (msgVersion), (MesTIFrom), (MesTITo))
            Print #1, "   </ScheduleTimeSeries>"
        End If
        col = col + 1
    Wend
    Print #1, "</ScheduleMessage>"
    Close #1
    MsgBox "Óñïåøíî å ãåíåðèðàí: " & File & msgVersion, vbQuestion, title

    ActiveSheet.Protect Password:="", AllowFormattingCells:=True, DrawingObjects:=False
    Application.ScreenUpdating = True
End Sub

Public Sub exportEssTimeSeries(col As Integer, msgVersion As Integer, MesTIFrom As String, MesTITo As String)
    Dim MMS As Worksheet
    Set MMS = Sheets("MMS")

    SendersTimeSeriesId = MMS.Cells(26, col).Value
    BusinessType = MMS.Cells(27, col).Value
    Product = MMS.Cells(28, col).Value
    ObjectAggregation = MMS.Cells(29, col).Value
    InArea = MMS.Cells(30, col).Value
    Incoding = MMS.Cells(31, col).Value
    OutArea = MMS.Cells(32, col).Value
    OutCoding = MMS.Cells(33, col).Value
    MeteringPoint = MMS.Cells(34, col).Value
    MeteringPointCoding = MMS.Cells(35, col).Value
    InParty = Trim(MMS.Cells(36, col).Value)
    InPartyCoding = MMS.Cells(37, col).Value
    OutParty = Trim(MMS.Cells(38, col).Value)
    OutPartyCoding = MMS.Cells(39, col).Value
    CapacityContractType = Trim(MMS.Cells(40, col).Value)
    CapacityAgreementIdentification = Trim(MMS.Cells(41, col).Value)
    MeasurementUnit = Trim(MMS.Cells(42, col).Value)
    Resolution = Trim(MMS.Cells(43, col).Value)

    Print #1, "      <"; "SendersTimeSeriesIdentification"; " v="""; SendersTimeSeriesId; """/>"
    Print #1, "      <"; "SendersTimeSeriesVersion"; " v="""; CStr(msgVersion); """/>"

    If IsEmpty(BusinessType) Then
            BusinessType = "A02"
    End If
    Print #1, "      <"; "BusinessType"; " v="""; BusinessType; """/>"

    If IsEmpty(Product) Then
            Product = "8716867000016"
    End If
    Print #1, "      <"; "Product"; " v="""; Product; """/>"

    If IsEmpty(ObjectAggregation) Then
            ObjectAggregation = "A01"
    End If
    Print #1, "      <"; "ObjectAggregation"; " v="""; ObjectAggregation; """/>"

    If Not (IsEmpty(InArea) Or InArea = "") Then
        If IsEmpty(Incoding) Then
                Incoding = "A01"
        End If
        Print #1, "      <InArea v="""; InArea; """ codingScheme="""; Incoding; """/>"
    End If

    If Not (IsEmpty(OutArea) Or OutArea = "") Then
        If IsEmpty(OutCoding) Then
                OutCoding = "A01"
        End If
        Print #1, "      <OutArea v="""; OutArea; """ codingScheme="""; OutCoding; """/>"
    End If

    If Not (IsEmpty(MeteringPoint) Or MeteringPoint = "") Then
        If IsEmpty(MeteringPointCoding) Then
                MeteringPointCoding = "A01"
        End If
        Print #1, "      <MeteringPointIdentification v="""; MeteringPoint; """ codingScheme="""; MeteringPointCoding; """/>"
    End If

    If Not (IsEmpty(InParty) Or InParty = "") Then
        If IsEmpty(InPartyCoding) Then
                InPartyCoding = "A01"
        End If
        Print #1, "      <InParty v="""; InParty; """ codingScheme="""; InPartyCoding; """/>"
    End If

    If Not (IsEmpty(OutParty) Or OutParty = "") Then
        If IsEmpty(OutPartyCoding) Then
                OutPartyCoding = "A01"
        End If
        Print #1, "      <OutParty v="""; OutParty; """ codingScheme="""; OutPartyCoding; """/>"
    End If

    If Not (IsEmpty(CapacityContractType) Or CapacityContractType = "") Then
            Print #1, "      <"; "CapacityContractType"; " v="""; CapacityContractType; """/>"
    End If

    If Not (IsEmpty(CapacityAgreementIdentification) Or CapacityAgreementIdentification = "") Then
            Print #1, "      <"; "CapacityAgreementIdentification"; " v="""; CapacityAgreementIdentification; """/>"
    End If

    If Not (IsEmpty(MeasurementUnit) Or MeasurementUnit = "") Then
            Print #1, "      <"; "MeasurementUnit"; " v="""; MeasurementUnit; """/>"
    End If

    Print #1, "      <Period>"
    Print #1, "      <TimeInterval v="""; MesTIFrom; "/"; MesTITo; """/>"

    If Not (IsEmpty(Resolution) Or Resolution = "") Then
            Print #1, "      <"; "Resolution"; " v="""; Resolution; """/>"
    End If

    Dim pos As Integer
    pos = 1
    qty = MMS.Cells(43 + pos, col)
    While Not IsEmpty(qty)
            Print #1, "         <Interval>"
            Print #1, "            <"; "Pos"; " v="""; CStr(pos); """/>"
            Print #1, "            <"; "Qty"; " v="""; Format(qty, "0.000"); """/>"
            Print #1, "         </Interval>"
            pos = pos + 1
            qty = MMS.Cells(43 + pos, col)
    Wend
    Print #1, "      </Period>"
End Sub
