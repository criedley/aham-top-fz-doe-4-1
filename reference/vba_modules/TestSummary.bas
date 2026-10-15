Attribute VB_Name = "TestSummary"

Public wkbData As Workbook
Sub PrintReport()
Attribute PrintReport.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim fname As String
    Dim wBook As Workbook
    Dim NumSheets As Integer
    Dim strPassword As String
    Dim Response1 As Variant
    
'    Dim wkbMain As Workbook

    strPassword = Sheets("x").Cells(1, 1) 'added in v1.20a
    
    AHAMSummary
        
    If Stable = True Then
        Sheets("Data Sheet").Activate
        Columns("A:B").AutoFit
        Columns("C:C").Select
        Selection.ColumnWidth = 16
        Range("A1").Select
        
'        wkbMain = ActiveWforkbook
        Sheets("Test Info").Activate
        ReportName = ""
        tsel = False
        isel = 1
        Do While tsel = False
            ' get string from the cell
            strtmp = Cells(isel, 1).Value
            ' compare
            strcs = StrComp(strtmp, "DATA FILENAME:")
            ' quick fix for room 34
            If strcs <> 0 Then
                strcs = StrComp(strtmp, "DATA_FILENAME:")
            End If
            If strcs = 0 Then
                strtmp = Cells(isel, 2).Value
                ReportName = Right(strtmp, (Len(strtmp) - InStrRev(strtmp, "\")))
                tsel = True
            Else
                isel = isel + 1
            End If
            If isel > 50 Then
                tsel = True
            End If
        Loop
        
        If Len(ReportName) > 4 Then
            fname = "Report_" & Left(ReportName, 5) & "DOE_SUMMARY___"
            'fname = fname & Mid(ReportName, 17, (Len(ReportName) - 4) - 16)
           fname = fname & Mid(ReportName, 17, (Len(ReportName) - 4) - 16) & "_DOE.xlsx"
        Else
            fname = "Report_"
        End If
        On Error GoTo n1
        Set wBook = Workbooks(fname & ".xlsx")
n1:
        If wBook Is Nothing Then

        Else
            wBook.Close
        End If
        
        fname = wbPath & "\" & fname 'added in v1.20a
        
        'Copy Picture of WattsTrace Chart from Main workbook to Report (v1.13)... modified in (v1.9)
        Set wBook = ActiveWorkbook
        If Not run100bool Then
        wkbThis.Sheets("T2 WattsTrace").Activate
        ActiveChart.ChartArea.Copy
        wBook.Unprotect
        wBook.Activate
        Sheets.Add before:=Sheets(1)
        ActiveSheet.Pictures.Paste.Select
        Selection.ShapeRange.IncrementLeft 192
        Selection.ShapeRange.Height = 430

        ActiveSheet.Shapes("Picture 1").Top = ActiveSheet.Rows(1).Top
        ActiveSheet.Name = "T2 WattsTrace"
        ActiveWindow.DisplayGridlines = False
        End If
        
        wkbThis.Sheets("T1 WattsTrace").Activate
        ActiveChart.ChartArea.Copy
        wBook.Unprotect
        wBook.Activate
        Sheets.Add before:=Sheets(1)
        ActiveSheet.Pictures.Paste.Select
        Selection.ShapeRange.IncrementLeft 192
        Selection.ShapeRange.Height = 430

        ActiveSheet.Shapes("Picture 1").Top = ActiveSheet.Rows(1).Top
        ActiveSheet.Name = "T1 WattsTrace"
        ActiveWindow.DisplayGridlines = False
        
        
        'Set wBook = ActiveWorkbook
        'wkbThis.Activate
        'Sheets("T1 WattsTrace").Select
        'ActiveChart.Activate
        'ActiveChart.ChartArea.Select
        'ActiveChart.CopyPicture Appearance:=xlPrinter, Size:=xlScreen, Format:=xlPicture
        
        'wBook.Activate
        'ActiveWorkbook.Unprotect
        'Charts.Add
        'ActiveChart.Name = "T1 WattsTrace"
        'ActiveSheet.Paste
        'ActiveChart.Deselect
        
        'wkbThis.Activate
        'Sheets("T2 WattsTrace").Select
        'ActiveChart.Activate
        'ActiveChart.ChartArea.Select
        'ActiveChart.CopyPicture Appearance:=xlPrinter, Size:=xlScreen, Format:=xlPicture
        
        'wBook.Activate
        'ActiveWorkbook.Unprotect
        'Charts.Add
        'ActiveChart.Name = "T2 WattsTrace"
        'ActiveSheet.Paste
        'ActiveChart.Deselect
        'ActiveWorkbook.Protect
        'Sheets("AHAM REPORT").Select
        'ActiveWorkbook.Save 'removed in v1.20a
        
        
SaveSheets:

    NumSheets = ActiveWorkbook.Sheets.Count
    For I = 1 To NumSheets
        Sheets(I).Protect Password:=strPassword
    Next I
    
    Application.DisplayAlerts = False 'added in v1.20a to remove dialogue box if file exists
    If Application.Version = "14.0" Then 'added in v1.20a to maintain compatibility between Excel 2010 and earlier versions
        'Response1 = MsgBox("Do you want to save report in Excel 2000/2003 format?", vbYesNo)
        'If Response1 <> 6 Then
            ActiveWorkbook.SaveAs Filename:=fname 'Save ONLY in Excel 2010 format for now (.xlsx)
        'Else:
            'ActiveWorkbook.SaveAs Filename:=fname, FileFormat:=xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
            'fname, FileFormat:=xlExcel14, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
            'fname, FileFormat:=xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        'End If
    Else:
        NewFN = Application.GetSaveAsFilename(fname, FileFilter:="Excell Files (*.xlsx), *.xlsx", Title:="Save Report as")

        If NewFN = False Then
            MsgBox "Report not saved"
        Else
            ActiveWorkbook.SaveAs NewFN
            'MsgBox "Report saved as " & NewFN
        End If
    End If
    
    Else:
        MsgBox "Stability not reached, report is not saved"
    End If
    
    Sheets("AHAM Report").Select
    Range("L5").Select
 
    'ActiveWorkbook.Close
End Sub

Sub AHAMSummary()
Attribute AHAMSummary.VB_ProcData.VB_Invoke_Func = " \n14"
Dim rngDC As Range
Dim rCell As Range
Dim DC As Integer
Dim ICEPass As Boolean
Dim VoltLimitUL As Double
Dim VoltLimitLL As Double
Dim tempRng As Range
Dim FFrealAVG As Double
Dim FZrealAVG As Double

    'Limits number of cycles shown to 100       Changed in v1.8
    If FullCycle > 100 Then FullCycle = 100
    If FullCycle >= UBound(TimeCycleOn) Then FullCycle = UBound(TimeCycleOn) - 1   'v1.9h
    
    
If run100bool Then
    For Count = 1 To FullCycle
    
        OnTime(Count) = Application.WorksheetFunction.Round(DateDiff("s", TimeCycleOn(Count), TimeCycleOn(Count + 1)) / 60, 2)
        OffTime(Count) = 0
    Next Count
Else:
    'For Count = 1 To stabcycle       Changed in v1.6
    For Count = 1 To FullCycle
        s = 1
        
        Do
            s = s + 1
        Loop Until Sheets("Events").Cells(s, 1) = TimeCycleOn(Count)
        
        OnTime(Count) = Application.WorksheetFunction.Round(DateDiff("s", Sheets("Events").Cells(s, 1), Sheets("Events").Cells(s + 1, 1)) / 60, 2)
        OffTime(Count) = Application.WorksheetFunction.Round(DateDiff("s", Sheets("Events").Cells(s + 1, 1), Sheets("Events").Cells(s + 2, 1)) / 60, 2)
        
    Next Count
End If
    
VoltLimitUL = 116
VoltLimitLL = 114

GraphWatts



'Section added in v1.9f  ************************
If FZOnly = False Then
    Set tempRng = Range(Cells(StartT1Chart, FFBW1col), Cells(EndT1Chart - 1, FFBW3col))
    FFrealAVG = Application.Average(tempRng)
End If
If FZBW4ch <> "NA" And FZBW5ch <> "NA" Then
    Set tempRng = Range(Cells(StartT1Chart, FZBW1col), Cells(EndT1Chart - 1, FZBW5col))
    FZrealAVG = Application.Average(tempRng)
ElseIf FZBW4ch <> "NA" Then
    Set tempRng = Range(Cells(StartT1Chart, FZBW1col), Cells(EndT1Chart - 1, FZBW4col))
    FZrealAVG = Application.Average(tempRng)
Else
    Set tempRng = Range(Cells(StartT1Chart, FZBW1col), Cells(EndT1Chart - 1, FZBW3col))
    FZrealAVG = Application.Average(tempRng)
End If
' *************************************************
    


'Arrange Worksheet order (v1.14c)
On Error Resume Next
Sheets("Data Sheet").Move before:=Sheets(1)
Sheets("Raw Data").Move After:=Sheets("Data Sheet")
Sheets("Events").Move After:=Sheets("Raw Data")
Sheets("Cycles").Move After:=Sheets("Events")
Sheets("Test Info").Move After:=Sheets("Cycles")
Sheets("Import File Info").Move After:=Sheets("Test Info")
On Error GoTo 0

'Find last populated line on Test Info sheet
Sheets("Test Info").Select
I = 1
Do Until Cells(I, 1) = "" And Cells(I + 1, 1) = "" And Cells(I + 2, 1) = "" 'Find Last Line.  Search for 3 consecutive blank rows.
    I = I + 1 'I = Last Row of Data + 1
Loop

Cells(I, 1) = "MACRO:"
Cells(I, 2) = Sheets("Data Sheet").Cells(30, 2)
Cells(I + 1, 1) = "Version:"
Cells(I + 1, 2) = Sheets("Data Sheet").Cells(31, 2)

I = I + 2

'Prompt User to enter Unit Serial Number
Range("A" & I) = "UNIT SERIAL NUMBER:"
Check1:
Range("B" & I).Select
Range("B" & I) = InputBox("Enter Unit Serial Number:")
Response1 = MsgBox("Is the S/N shown in Cell B" & I & " correct?", vbYesNo)
If Response1 <> 6 Then GoTo Check1
If Range("B" & I).Value = "" Then GoTo SkipInputs

'Prompt User to select Compressor Run-In Time from Test Info Sheet
I = I + 1
Range("A" & I) = "Compressor Run-In Time:"
Columns(1).AutoFit
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=">24 hours,>12 hours,NA"
End With
UserForm1.Show
UserForm1.caption = "Compressor Run-In Time"
UserForm1.TextBox1 = "Select Compressor Run-In Time from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Anti-Sweat Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Anti-Sweat Heaters:"
Columns(1).AutoFit
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="ON,OFF,NA"
End With
UserForm1.Show
UserForm1.caption = "Anti-Sweat Heaters"
UserForm1.TextBox1 = "Select Anti-Sweat Mode from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Meat Pan Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Meat Pan:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Cold, Colder, Coldest Available, Low, Open Notch, Closed Notch, Meat,NA"
End With
UserForm1.Show
UserForm1.caption = "Meat Pan"
UserForm1.TextBox1 = "Select Meat Pan Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Veg Pan Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Veg Pan:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Coldest Available, Low, Open Notch, Closed Notch, Fruits, Vegetables,NA"
End With
UserForm1.Show
UserForm1.caption = "Veg Pan"
UserForm1.TextBox1 = "Select Veg Pan Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Veg Pan Right Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Veg Pan Right:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Coldest Available, Low, Open Notch, Closed Notch, Fruits, Vegetables,NA"
End With
UserForm1.Show
UserForm1.caption = "Veg Pan Right"
UserForm1.TextBox1 = "Select Veg Pan Right Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Veg Pan Left Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Veg Pan Left:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Coldest Available, Low, Open Notch, Closed Notch, Fruits, Vegetables,NA"
End With
UserForm1.Show
UserForm1.caption = "Veg Pan Left"
UserForm1.TextBox1 = "Select Veg Pan Left Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Veg Pan Left Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Veg Pan Up:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Coldest Available, Low, Open Notch, Closed Notch, Fruits, Vegetables,NA"
End With
UserForm1.Show
UserForm1.caption = "Veg Pan Up"
UserForm1.TextBox1 = "Select Veg Pan Up Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Veg Pan Left Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Veg Pan Middle:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Coldest Available, Low, Open Notch, Closed Notch, Fruits, Vegetables,NA"
End With
UserForm1.Show
UserForm1.caption = "Veg Pan Middle"
UserForm1.TextBox1 = "Select Veg Pan Middle Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Veg Pan Left Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Veg Pan Low:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Coldest Available, Low, Open Notch, Closed Notch, Fruits, Vegetables,NA"
End With
UserForm1.Show
UserForm1.caption = "Veg Pan Low"
UserForm1.TextBox1 = "Select Veg Pan Low Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Deli Pan Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Deli Pan:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Open Notch, Closed Notch, Coldest Available, Citrus, Beverage, Cheese, Produce, Meat, NA"
End With
UserForm1.Show
UserForm1.caption = "Deli Pan"
UserForm1.TextBox1 = "Select Deli Pan Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Snack Pan Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Snack Pan:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Open Notch,Coldest,Coldest Available,NA"
End With
UserForm1.Show
UserForm1.caption = "Snack Pan"
UserForm1.TextBox1 = "Select Snack Pan Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Custom Cool Pan Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Custom Cool Pan:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Meat,Produce,Citrus,Off,NA"
End With
UserForm1.Show
UserForm1.caption = "Custom Cool Pan"
UserForm1.TextBox1 = "Select Custom Cool Pan Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Ice Maker Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Ice Maker:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="ON,OFF,ON-HARVEST DISABLED,NA"
End With
UserForm1.Show
UserForm1.caption = "Ice Maker"
UserForm1.TextBox1 = "Select Ice Maker Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

'Prompt User to select Wi-Fi Mode from Test Info Sheet
I = I + 1
Range("A" & I) = "Wi-Fi:"
Range("B" & I).Select
Range("B" & I).Clear
With Selection.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="NA,ON - Not Communicating"
End With
UserForm1.Show
UserForm1.caption = "Wi-Fi"
UserForm1.TextBox1 = "Select Wi-Fi Type from Cell B" & I
Do Until Range("B" & I) <> ""
    DoEvents
Loop

SkipInputs:

UserForm1.Hide
Range("A" & I + 1).Select
'******************************************************
Get_Computer_Name
Get_User_Name
    Set wkbData = ActiveWorkbook
    Application.ScreenUpdating = False
    Workbooks.Add xlWBATWorksheet
    Set wkbReport = ActiveWorkbook
    wkbReport.Sheets(1).Name = "AHAM Report"
    offset = 1
    ' continue collecting data
    Cells((1 + offset), 2).Value = "Test Duration"
    Cells((1 + offset), 3).Value = Application.WorksheetFunction.Round(TestDuration, 1)
    Cells((1 + offset), 4).Value = "min"
    Cells((1 + offset), 5).Value = "(" & Application.WorksheetFunction.Round(TestDuration / 60, 1) & "hrs)"
    Cells((2 + offset), 2).Value = "Energy"
    Cells((2 + offset), 3).Value = Application.WorksheetFunction.Round(ET, 3)
    Cells((2 + offset), 4).Value = "kWh/day"

    Cells((3 + offset), 2).Value = "EP1"
    Cells((3 + offset), 3).Value = EP1
    Cells((3 + offset), 4).Value = "kWh"
    Cells((4 + offset), 2).Value = "T1"
    Cells((4 + offset), 3).Value = Application.WorksheetFunction.Round(t1, 1)
    Cells((4 + offset), 4).Value = "min"
    Cells((5 + offset), 2).Value = "T1 Period"
    Cells((5 + offset), 3).Value = T1TimeOn
    Cells((5 + offset), 4).Value = "To"
    Cells((5 + offset), 5).Value = TimeCycleOn(stabcycle + 1)
    If ChooseCycleBool = True Then
        Cells((5 + offset), 5).Value = TimeCycleOn(CycleAmount2 + 1)
    End If
        Cells((6 + offset), 2).Value = "EP2"
    Cells((6 + offset), 3).Value = EP2
    Cells((6 + offset), 4).Value = "kWh"
    Cells((7 + offset), 2).Value = "T2"
    Cells((7 + offset), 3).Value = Application.WorksheetFunction.Round(T2, 1)
    Cells((7 + offset), 4).Value = "min"
    Cells((8 + offset), 2).Value = "T2 Period"
    Cells((8 + offset), 3).Value = T2TimeOn
    Cells((8 + offset), 4).Value = "To"
    Cells((8 + offset), 5).Value = T2TimeOff
    Cells((9 + offset), 2).Value = "FZ CT"
    Cells((9 + offset), 3).Value = CT
    Cells((10 + offset), 2).Value = "FF Avg"
    If FZOnly = False Then
        Cells((10 + offset), 3).Value = Application.WorksheetFunction.Round(FFrealAVG, 2)
    Else:
        Cells((10 + offset), 3).Value = "NA"
    End If
    Cells((10 + offset), 4).Value = "°F"
    Cells((11 + offset), 2).Value = "FZ Avg"
    Cells((11 + offset), 3).Value = Application.WorksheetFunction.Round(FZrealAVG, 2)
    Cells((11 + offset), 4).Value = "°F"
    Cells((12 + offset), 2).Value = "Ambient"
    Cells((13 + offset), 2).Value = "Gradient"
    Cells((14 + offset), 2).Value = "ICE BW"
    Cells((15 + offset), 2).Value = "K factor"
    Cells((15 + offset), 3).Value = FZOnlyKfactor
    Cells((16 + offset), 2).Value = "Energy Corr."
    Cells((16 + offset), 3).Value = Application.WorksheetFunction.Round(ETcorr, 3)
    Cells((16 + offset), 4).Value = "kWh/day"
    Cells((17 + offset), 2).Value = "LSOC Avg"
    Cells((17 + offset), 3).Value = LSOCTotAvg
    Cells((17 + offset), 4).Value = "°F"
    Cells((18 + offset), 2).Value = "RSOC Avg"
    Cells((18 + offset), 3).Value = RSOCTotAvg
    Cells((18 + offset), 4).Value = "°F"
    
    ' align cells
    With Range(Cells(1 + offset, 1), Cells(18 + offset, 2)).Interior
        .ColorIndex = 49
        .Pattern = xlSolid
    End With
    With Range(Cells(1 + offset, 3), Cells(18 + offset, 6)).Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With
    

    Range(Cells(1 + offset, 1), Cells(18 + offset, 2)).Font.ColorIndex = 2
    Range(Cells(1 + offset, 1), Cells(18 + offset, 2)).Font.Bold = True
    Range(Cells((1 + offset), 2), Cells((18 + offset), 2)).HorizontalAlignment = xlRight
    Range(Cells((1 + offset), 4), Cells((18 + offset), 4)).HorizontalAlignment = xlLeft
    Set rngDC = Range(Cells(1 + offset, 1), Cells(18 + offset, 6))
    rngDC.Borders(xlDiagonalDown).LineStyle = xlNone
    rngDC.Borders(xlDiagonalUp).LineStyle = xlNone
    With rngDC.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With

    For I = 1 To 17
        Set rngDC = Range(Cells(offset + I, 3), Cells(offset + I, 6))
        With rngDC.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next

    'Add Hi-Med & Med-Lo Gradients
    Cells(2, 11).Value = "Left Gradient Lo-Mid  "
    Cells(3, 11).Value = "Left Gradient Mid-Hi  "
    Cells(4, 11).Value = "Right Gradient Lo-Mid  "
    Cells(5, 11).Value = "Right Gradient Mid-Hi  "
    Cells(2, 12).Value = MLgradLeftMax
    Cells(3, 12).Value = HMgradLeftMax
    Cells(4, 12).Value = MLgradRightMax
    Cells(5, 12).Value = HMgradRightMax

    With Range(Cells(2, 8), Cells(5, 11)).Interior
        .ColorIndex = 49
        .Pattern = xlSolid
    End With
    With Range(Cells(2, 12), Cells(5, 12)).Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With
    
    Range(Cells(2, 8), Cells(5, 11)).Font.ColorIndex = 2
    Range(Cells(2, 8), Cells(5, 11)).Font.Bold = True
    Range(Cells(2, 8), Cells(5, 11)).HorizontalAlignment = xlRight
    Range(Cells(2, 12), Cells(5, 12)).HorizontalAlignment = xlCenter
    Set rngDC = Range(Cells(2, 8), Cells(5, 12))
    rngDC.Borders(xlDiagonalDown).LineStyle = xlNone
    rngDC.Borders(xlDiagonalUp).LineStyle = xlNone
    With rngDC.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With

    For I = 2 To 4
        Set rngDC = Range(Cells(I, 8), Cells(I, 12))
        With rngDC.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next






    offset = offset + 19
    If cyclic = True Then
        Cells((1 + offset), 1).Value = "Cycle"
    Else
        Cells((1 + offset), 1).Value = "Hour"
    End If
    Cells((1 + offset), 2).Value = "FF"
    Cells((2 + offset), 2).Value = "BW 1"
    Cells((1 + offset), 3).Value = "FF"
    Cells((2 + offset), 3).Value = "BW 2"
    Cells((1 + offset), 4).Value = "FF"
    Cells((2 + offset), 4).Value = "BW 3"
    Cells((1 + offset), 5).Value = "FZ"
    Cells((2 + offset), 5).Value = "BW 1"
    Cells((1 + offset), 6).Value = "FZ"
    Cells((2 + offset), 6).Value = "BW 2"
    Cells((1 + offset), 7).Value = "FZ"
    Cells((2 + offset), 7).Value = "BW 3"
    Cells((1 + offset), 8).Value = "FZ"
    Cells((2 + offset), 8).Value = "BW 4"
    Cells((1 + offset), 9).Value = "FZ"
    Cells((2 + offset), 9).Value = "BW 5"
    Cells((1 + offset), 10).Value = "ICE"
    Cells((2 + offset), 10).Value = "BW"
    Cells((1 + offset), 11).Value = "Volts"
    Cells((2 + offset), 11).Value = "Avg"
    Cells((1 + offset), 12).Value = "Volts"
    Cells((2 + offset), 12).Value = "Min"
    Cells((1 + offset), 13).Value = "Volts"
    Cells((2 + offset), 13).Value = "Max"
       
    Cells((1 + offset), 14).Value = "LSOC"
    Cells((2 + offset), 14).Value = "Avg"
    Cells((1 + offset), 15).Value = "RSOC"
    Cells((2 + offset), 15).Value = "Avg"
    Cells((1 + offset), 16).Value = "Left "
    Cells((2 + offset), 16).Value = "Gradient"
    Cells((1 + offset), 17).Value = "Right "
    Cells((2 + offset), 17).Value = "Gradient"
    Cells((1 + offset), 18).Value = "Cycle "
    Cells((2 + offset), 18).Value = "Duration"
    Cells((1 + offset), 19).Value = "On"
    Cells((2 + offset), 19).Value = "Time"
    Cells((1 + offset), 20).Value = "Off"
    Cells((2 + offset), 20).Value = "Time"
    Cells((1 + offset), 21).Value = "% "
    Cells((2 + offset), 21).Value = "Runtime"
    Cells((1 + offset), 22).Value = "WHr "
    Cells((2 + offset), 22).Value = "Consumption"
    Cells((1 + offset), 23).Value = "Start of "
    Cells((2 + offset), 23).Value = "Cycle"
    Cells((1 + offset), 24).Value = "End of "
    Cells((2 + offset), 24).Value = "Cycle"

    ' blue title
    With Range(Cells(1 + offset, 1), Cells(2 + offset, 24)).Interior
        .ColorIndex = 49
        .Pattern = xlSolid
    End With

    ' centered title
    Range(Cells(1 + offset, 1), Cells(2 + offset, 24)).Font.ColorIndex = 2
    Range(Cells(1 + offset, 1), Cells(2 + offset, 24)).Font.Bold = True
    Range(Cells(1 + offset, 1), Cells(2 + offset, 24)).HorizontalAlignment = xlCenter

    ' blue side (cycles)
    With Range(Cells(1 + offset, 1), Cells(2 + offset + FullCycle, 1)).Interior
        .ColorIndex = 49
        .Pattern = xlSolid
    End With

    ' centered side
    Range(Cells(1 + offset, 1), Cells(4 + offset + FullCycle, 1)).Font.ColorIndex = 2
    Range(Cells(1 + offset, 1), Cells(4 + offset + FullCycle, 1)).Font.Bold = True
    Range(Cells(1 + offset, 1), Cells(4 + offset + FullCycle, 1)).HorizontalAlignment = xlCenter

    offset = offset + 3
    ' paint background before displaying the result
    With Range(Cells(offset, 2), Cells(offset + FullCycle - 1, 24)).Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With

    ' set thin borders
    For I = -1 To FullCycle - 1
        With Range(Cells(offset + I, 1), Cells(offset + I, 24)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next I

offset = offset - 1


If ChooseCycleBool = True Then
    For I = 0 To CycleAmount2 - CycleAmount
        Range(Cells(offset + CycleAmount + I, 2), Cells(offset + CycleAmount + I, 24)).Interior.Color = RGB(255, 255, 110)
    Next I
End If
    
    
    'For I = 1 To stabcycle       Changed in v1.6
    For I = 1 To FullCycle
        Cells((offset + I), 1).Value = I
        Cells((offset + I), 2).Value = FFBW1(I)
        Cells((offset + I), 3).Value = FFBW2(I)
        Cells((offset + I), 4).Value = FFBW3(I)
        Cells((offset + I), 5).Value = FZBW1(I)
        Cells((offset + I), 6).Value = FZBW2(I)
        Cells((offset + I), 7).Value = FZBW3(I)
        Cells((offset + I), 8).Value = FZBW4(I)
        Cells((offset + I), 9).Value = FZBW5(I)
        Cells((offset + I), 10).Value = ICEBW(I)
    
        Cells((offset + I), 11).Value = VoltsAv(I)
        If Not (VoltsAv(I) > VoltLimitLL And VoltsAv(I) < VoltLimitUL) Then
            Cells((offset + I), 11).Interior.ColorIndex = 45
        End If
                
        Cells((offset + I), 12).Value = VoltsMin(I)
        If Not (VoltsMin(I) > VoltLimitLL And VoltsMin(I) < VoltLimitUL) Then
            Cells((offset + I), 12).Interior.ColorIndex = 45
        End If
        
        
        Cells((offset + I), 13).Value = VoltsMax(I)
        If Not (VoltsMax(I) > VoltLimitLL And VoltsMax(I) < VoltLimitUL) Then
            Cells((offset + I), 13).Interior.ColorIndex = 45
        End If
        
        Cells((offset + I), 14).Value = LSOCAvg(I)
        If Not (LSOCAvg(I) > 88.95 And LSOCAvg(I) < 91.05) Then
            If I <= stabcycle Then
                AmbientFLG = False
            End If
            Cells((offset + I), 14).Interior.ColorIndex = 45
        End If
        Cells((offset + I), 15).Value = RSOCAvg(I)
        If Not (RSOCAvg(I) > 88.95 And RSOCAvg(I) < 91.05) Then
            If I <= stabcycle Then
                AmbientFLG = False
            End If
            Cells((offset + I), 15).Interior.ColorIndex = 45
        End If
        Cells((offset + I), 16).Value = LGrad(I)
        If Abs(LGrad(I)) > GradientLimit Then
            If I <= stabcycle Then
                GradientFLG = False
            End If
            Cells((offset + I), 16).Interior.ColorIndex = 45
        End If
        Cells((offset + I), 17).Value = RGrad(I)
        If Abs(RGrad(I)) > GradientLimit Then
            If I <= stabcycle Then
                GradientFLG = False
            End If
            Cells((offset + I), 17).Interior.ColorIndex = 45
        End If
        If cyclic = True Then
            If I < cycleCount Then
                Cells((offset + I), 18).Value = Application.WorksheetFunction.Round((DateDiff("s", TimeCycleOn(I), TimeCycleOn(I + 1)) / 60), 2)
                Cells((offset + I), 19).Value = OnTime(I)
                Cells((offset + I), 20).Value = OffTime(I)
                Cells((offset + I), 21).Value = 0
                If (OnTime(I) + OffTime(I)) <> 0 Then       'Added in v1.9j
                    Cells((offset + I), 21).Value = (OnTime(I) / (OnTime(I) + OffTime(I)))
                Else
                    Cells((offset + I), 20).Value = Application.WorksheetFunction.Round((DateDiff("s", TimeCycleOn(I), TimeCycleOn(I + 1)) / 60), 2)
                End If
            Else
                Cells((offset + I), 18).Value = Application.WorksheetFunction.Round((DateDiff("s", TimeCycleOn(I), LastScan) / 60), 2)
                Cells((offset + I), 19).Value = OnTime(I)
                Cells((offset + I), 20).Value = OffTime(I)
                Cells((offset + I), 21).Value = 0
                If (OnTime(I) + OffTime(I)) <> 0 Then       'Added in v1.9j
                    Cells((offset + I), 21).Value = (OnTime(I) / (OnTime(I) + OffTime(I)))
                Else
                    Cells((offset + I), 20).Value = Application.WorksheetFunction.Round((DateDiff("s", TimeCycleOn(I), LastScan) / 60), 2)
                End If
            End If
        Else
            If I < cycleCount Then
                Cells((offset + I), 18).Value = Application.WorksheetFunction.Round((DateDiff("s", TimeCycleOn(I), TimeCycleOn(I + 1)) / 60), 2)
                Cells((offset + I), 19).Value = OnTime(I)
                Cells((offset + I), 20).Value = OffTime(I)
                Cells((offset + I), 21).Value = 1
            Else
                Cells((offset + I), 18).Value = Application.WorksheetFunction.Round((DateDiff("s", TimeCycleOn(I), LastScan) / 60), 2)
                Cells((offset + I), 19).Value = OnTime(I)
                Cells((offset + I), 20).Value = OffTime(I)
                Cells((offset + I), 21).Value = 1
            End If
        End If
        If run100bool Then Range(Cells(offset + I, 18), Cells(offset + I, 20)).NumberFormat = "0"
        If I < cycleCount Then
            Cells(offset + I, 22).Value = WattHoursCount(I + 1) - WattHoursCount(I)
        Else
            Cells(offset + I, 22).Value = LastWattHourCount - WattHoursCount(I)
        End If
        
        With Range(Cells(offset + I, 1), Cells(offset + 1, 22)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        
        'Display Start/End Times of Cycles
        Cells((offset + I), 23).Value = TimeCycleOn(I)
        Cells((offset + I), 24).Value = TimeCycleOn(I + 1)
    Next
    

    
    datastart = offset + I
    
    Cells(9 + offset - 18, 3).Value = AmbientFLG
    Cells(10 + offset - 18, 3).Value = GradientFLG
    
    Range(Cells(offset, 21), Cells(offset + FullCycle, 21)).NumberFormat = "0.00%"
    
    Cells(offset + stabcycle, 25).Value = "stable"
    With Cells(offset + stabcycle, 25).Interior
        .ColorIndex = 4
        .Pattern = xlSolid
    End With
    Cells(offset + stabcycle, 25).Font.Bold = True
    With Cells(offset + stabcycle, 25)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With

    Set rngDC = Range(Cells(offset + stabcycle, 1), Cells(offset + stabcycle, 25))
    rngDC.Borders(xlDiagonalDown).LineStyle = xlNone
    rngDC.Borders(xlDiagonalUp).LineStyle = xlNone
    With rngDC.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With



    Range(Cells(offset, 2), Cells(offset + FullCycle, 24)).HorizontalAlignment = xlCenter
    Set rngDC = Range(Cells(offset - 1, 1), Cells(offset + FullCycle, 24))
    rngDC.Borders(xlDiagonalDown).LineStyle = xlNone
    rngDC.Borders(xlDiagonalUp).LineStyle = xlNone
    With rngDC.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rngDC.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With

    rngDC.Columns.AutoFit
    Columns(12).ColumnWidth = 6

    For I = 1 To 24
        Set rngDC = Range(Cells(offset - 1, I), Cells(offset + FullCycle, I))
        With rngDC.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next
    
    'Run check for ICE BW (<=32F)==================================================
    'Assume ICE BW is always Column 10
    C = 10: R = 1
    Do Until Cells(R, C) = "ICE"
        If R > 27 Then 'Allow search for 20 lines before Avg Data starts
            MsgBox ("ICE BW not found.  Macro error. Contact Programmer."), vbOKOnly
            End
        End If
        R = R + 1
    Loop
    R = R + 2 'Start of Averaged Data
    'Check each line of Averaged Data for Temps <=32F
    ICEPass = True
    Do Until Cells(R, 1) = ""
        If Cells(R, C) > 32 Then
            Cells(R, C).Interior.ColorIndex = 45 'Color Cell if >32F
            ICEPass = False
        End If
        R = R + 1
    Loop
    If ICEBW(1) <> 0 Then
        Cells(11 + offset - 18, 3) = ICEPass
        If ICEPass = False Then
        
        End If
    Else:
        Cells(11 + offset - 18, 3) = "NA"
    End If
    
    'If Ambient or Gradient or ICE BW Flag is False, the lengthen column C by 1 character.
    If AmbientFLG = False Or GradientFLG = False Or ICEPass = False Then
        Columns("C").ColumnWidth = 6
    End If

'Print average row
'Removed per request, but left in code just in case it is needed later
    'avgrow = datastart
    'PrintAverage

If run100bool Then GoTo skip100
'Print T2 Data  *********************************************************************************************
Cells(datastart + 2, 1).Value = "T2 Data"
Range(Cells(datastart + 2, 1), Cells(datastart + 2, 1)).Font.Bold = True
offset = datastart + 2
If cyclic = True Then
        Cells((1 + offset), 1).Value = "Cycle"
    Else
        Cells((1 + offset), 1).Value = "Hour"
    End If
    Cells((1 + offset), 2).Value = "FF"
    Cells((2 + offset), 2).Value = "BW 1"
    Cells((1 + offset), 3).Value = "FF"
    Cells((2 + offset), 3).Value = "BW 2"
    Cells((1 + offset), 4).Value = "FF"
    Cells((2 + offset), 4).Value = "BW 3"
    Cells((1 + offset), 5).Value = "FZ"
    Cells((2 + offset), 5).Value = "BW 1"
    Cells((1 + offset), 6).Value = "FZ"
    Cells((2 + offset), 6).Value = "BW 2"
    Cells((1 + offset), 7).Value = "FZ"
    Cells((2 + offset), 7).Value = "BW 3"
    Cells((1 + offset), 8).Value = "FZ"
    Cells((2 + offset), 8).Value = "BW 4"
    Cells((1 + offset), 9).Value = "FZ"
    Cells((2 + offset), 9).Value = "BW 5"
    Cells((1 + offset), 10).Value = "ICE"
    Cells((2 + offset), 10).Value = "BW"
    Cells((1 + offset), 11).Value = "Volts"
    Cells((2 + offset), 11).Value = "Avg"
    Cells((1 + offset), 12).Value = "Volts"
    Cells((2 + offset), 12).Value = "Min"
    Cells((1 + offset), 13).Value = "Volts"
    Cells((2 + offset), 13).Value = "Max"
       
    Cells((1 + offset), 14).Value = "LSOC"
    Cells((2 + offset), 14).Value = "Avg"
    Cells((1 + offset), 15).Value = "RSOC"
    Cells((2 + offset), 15).Value = "Avg"
    Cells((1 + offset), 16).Value = "Left "
    Cells((2 + offset), 16).Value = "Gradient"
    Cells((1 + offset), 17).Value = "Right "
    Cells((2 + offset), 17).Value = "Gradient"
    Cells((1 + offset), 18).Value = "Cycle "
    Cells((2 + offset), 18).Value = "Duration"
    Cells((1 + offset), 19).Value = "On"
    Cells((2 + offset), 19).Value = "Time"
    Cells((1 + offset), 20).Value = "Off"
    Cells((2 + offset), 20).Value = "Time"
    Cells((1 + offset), 21).Value = "% "
    Cells((2 + offset), 21).Value = "Runtime"
    Cells((1 + offset), 22).Value = "WHr "
    Cells((2 + offset), 22).Value = "Consumption"
    Cells((1 + offset), 23).Value = "Start of "
    Cells((2 + offset), 23).Value = "Cycle"
    Cells((1 + offset), 24).Value = "End of "
    Cells((2 + offset), 24).Value = "Cycle"

    t2cycs = 5 + T2Pre + T2Post

    ' blue title
    With Range(Cells(1 + offset, 1), Cells(2 + offset, 24)).Interior
        .ColorIndex = 49
        .Pattern = xlSolid
    End With

    ' centered title
    Range(Cells(1 + offset, 1), Cells(2 + offset, 24)).Font.ColorIndex = 2
    Range(Cells(1 + offset, 1), Cells(2 + offset, 24)).Font.Bold = True
    Range(Cells(1 + offset, 1), Cells(2 + offset, 24)).HorizontalAlignment = xlCenter

    ' blue side (cycles)
    With Range(Cells(1 + offset, 1), Cells(offset + t2cycs + 2, 1)).Interior
        .ColorIndex = 49
        .Pattern = xlSolid
    End With

    ' centered side
    Range(Cells(1 + offset, 1), Cells(offset + t2cycs + 2, 1)).Font.ColorIndex = 2
    Range(Cells(1 + offset, 1), Cells(offset + t2cycs + 2, 1)).Font.Bold = True
    Range(Cells(1 + offset, 1), Cells(offset + t2cycs + 2, 1)).HorizontalAlignment = xlCenter

    
    ' paint background before displaying the result
    With Range(Cells(offset + 3, 2), Cells(offset + t2cycs + 2, 24)).Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With

    ' set thin borders
    For I = 1 To 2
        With Range(Cells(offset + I + 1, 1), Cells(offset + I + 1, 24)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next I
    
    
    
    offset = offset + 2
    For I = 1 To t2cycs
        If I = 1 Then
            Cells(offset + I, 1) = "A Cycle"
            xcol = offset + I
        ElseIf I = 2 Then
            Cells(offset + I + T2Pre, 1) = "Precool"
            xcol = offset + I + T2Pre
        ElseIf I = 3 Then
            Cells(offset + I + T2Pre, 1) = "Defrost"
            xcol = offset + I + T2Pre
        ElseIf I = 4 Then
            Cells(offset + I + T2Pre, 1) = "Recovery"
            xcol = offset + I + T2Pre
        ElseIf I = 5 Then
            Cells(offset + I + T2Pre + T2Post, 1) = "B Cycle"
            xcol = offset + I + T2Pre + T2Post
        ElseIf I < (5 + T2Pre + 1) Then
            Cells(offset + I - 4, 1) = "Pre " & (I - 5)
            xcol = offset + I - 4
        Else
            Cells(offset + I - 1, 1) = "Post " & (I - T2Pre - 5)
            xcol = offset + I - 1
        End If
        
        Cells(xcol, 2) = T2FF1(I)
        Cells(xcol, 3) = T2FF2(I)
        Cells(xcol, 4) = T2FF3(I)
        Cells(xcol, 5) = T2FZ1(I)
        Cells(xcol, 6) = T2FZ2(I)
        Cells(xcol, 7) = T2FZ3(I)
        Cells(xcol, 8) = T2FZ4(I)
        Cells(xcol, 9) = T2FZ5(I)
        Cells(xcol, 10) = T2IB(I)
        Cells(xcol, 11) = T2Volts(I)
        If Not (T2Volts(I) > VoltLimitLL And T2Volts(I) < VoltLimitUL) Then
            Cells((xcol), 11).Interior.ColorIndex = 45
        End If
            
        Cells(xcol, 12) = T2VoltsMin(I)
        If Not (T2VoltsMin(I) > VoltLimitLL And T2VoltsMin(I) < VoltLimitUL) Then
            Cells((xcol), 12).Interior.ColorIndex = 45
        End If
        
        Cells(xcol, 13) = T2VoltsMax(I)
        If Not (T2VoltsMax(I) > VoltLimitLL And T2VoltsMax(I) < VoltLimitUL) Then
            Cells(xcol, 13).Interior.ColorIndex = 45
        End If
        
        Cells(xcol, 14) = T2LsocAvg(I)
        If Not (T2LsocAvg(I) > 88.95 And T2LsocAvg(I) < 91.05) Then
            If I <= stabcycle Then
                AmbientFLG = False
            End If
            Cells(xcol, 14).Interior.ColorIndex = 45
        End If
        
        Cells(xcol, 15) = T2RsocAvg(I)
        If Not (T2RsocAvg(I) > 88.95 And T2RsocAvg(I) < 91.05) Then
            If I <= stabcycle Then
                AmbientFLG = False
            End If
            Cells(xcol, 15).Interior.ColorIndex = 45
        End If
        
        Cells(xcol, 16) = T2LGrad(I)
        If Abs(T2LGrad(I)) > GradientLimit Then
            If I <= stabcycle Then
                GradientFLG = False
            End If
            Cells(xcol, 16).Interior.ColorIndex = 45
        End If
        
        Cells(xcol, 17) = T2Rgrad(I)
        If Abs(T2Rgrad(I)) > GradientLimit Then
            If I <= stabcycle Then
                GradientFLG = False
            End If
            Cells(xcol, 17).Interior.ColorIndex = 45
        End If
        

          If cyclic = True Then
                    Cells(xcol, 18).Value = T2Run(I)
                    Cells(xcol, 19).Value = Abs(T2OnTime(I))
                    Cells(xcol, 20).Value = Abs(T2OffTime(I))
                    Cells(xcol, 21).Value = (Abs(T2OnTime(I)) / (Abs(T2OnTime(I)) + Abs(T2OffTime(I))))
          Else
              Cells(xcol, 18).Value = Application.WorksheetFunction.Round((DateDiff("s", T2On(I), T2On(I + 1)) / 60), 2)
              Cells(xcol, 19).Value = Abs(T2OnTime(I))
              Cells(xcol, 20).Value = Abs(T2OffTime(I))
              Cells(xcol, 21).Value = 1
          End If
  

        
        Range(Cells(xcol, 21), Cells(xcol, 21)).NumberFormat = "0.00%"
        Cells(xcol, 22).Value = T2WH(I)
        Cells(xcol, 23).Value = T2On(I)
        Cells(xcol, 24).Value = T2Off(I)
        
        
        
    
        With Range(Cells(xcol, 1), Cells(xcol, 24)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With Range(Cells(xcol, 1), Cells(xcol, 24))
            .HorizontalAlignment = xlCenter
        End With
    Next I
    
    For x = 1 To 24
        Set rngDC = Range(Cells(offset - 1, x), Cells(xcol + 1, x))
        With rngDC.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next
    

'Print T2 average row
'Removed  per request but left in code, in case needed later
    'avgrow = offset + t2cycs + 1
    'PrintAverage
    
skip100:
    
    Columns("C:C").AutoFit
    Columns("E:E").AutoFit
    Columns("U:V").AutoFit
    
    wkbThis.Sheets("Test Info").Copy before:=Sheets("AHAM Report")
    Range(Sheets("Test Info").Cells(1, 1), Sheets("Test Info").Cells(50, 2)).Columns.AutoFit
    Range(Sheets("Test Info").Cells(1, 2), Sheets("Test Info").Cells(50, 2)).HorizontalAlignment = xlLeft
    wkbThis.Sheets("Events").Copy before:=Sheets("AHAM Report")
    Range(Sheets("Events").Cells(1, 1), Sheets("Events").Cells(1000, 9)).Columns.AutoFit
    Range(Sheets("Events").Cells(1, 1), Sheets("Events").Cells(1000, 9)).HorizontalAlignment = xlCenter
    wkbThis.Sheets("Import File Info").Copy before:=Sheets("AHAM Report")
    'wkbThis.Sheets("Data Sheet").Copy Before:=Sheets("AHAM Report")
    'Sheets("Data Sheet").Activate
    'rngDC =
    'rngDC.Select
    
    Sheets.Add Type:="Worksheet"
    With ActiveSheet
        .Name = "Data Sheet"
    End With
    wkbData.Activate
    Sheets("Data Sheet").Activate
    
    Range(Cells(1, 1), Cells(21, 7)).Copy Destination:=wkbReport.Sheets("Data Sheet").Cells(1, 1)
    A = Range("G3").Value 'added in v1.20a
    wkbReport.Sheets("Data Sheet").Range("G3") = A 'added in v1.20a
    wkbReport.Activate
    
    Range(Sheets("Data Sheet").Cells(1, 1), Sheets("Data Sheet").Cells(21, 9)).Select
    With Selection.Validation
        .Delete
    End With
    Cells(1, 1).Select
    Sheets("Import File Info").Activate
    Cells(1, 1).Value = "Username: "
    Cells(1, 2).Value = UserName
    Cells(2, 1).Value = "Computer Name: "
    Cells(2, 2).Value = ComputerName
    Cells(3, 2) = wkbData.Sheets("Info").Cells(7, 2)
    Sheets("AHAM Report").Activate
    
    Columns("A:A").EntireColumn.AutoFit
    Columns("W:W").EntireColumn.AutoFit
    Columns("X:X").EntireColumn.AutoFit
    
    wkbReport.Protect
    Cells(1, 1).Select
    
End Sub

