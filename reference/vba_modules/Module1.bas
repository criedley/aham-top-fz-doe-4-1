Attribute VB_Name = "Module1"
Option Explicit
' GetCycles()
' ClearFields()
' CommandButton1_Click()
' ChooseCycleAmount()
' SheetExists()
' ImportData()
Public watts As Integer
Public wbPath As String 'added in v1.20a


Sub ClearFields()
Attribute ClearFields.VB_ProcData.VB_Invoke_Func = " \n14"
Dim oldStatusBar As Boolean
oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    Sheets("Raw Data").Activate
    Cells.Select
    Selection.ClearContents
    Cells(1, 1).Select
    Sheets("Import File Info").Activate
    Cells.Select
    Selection.ClearContents
    Cells(1, 1).Select
    Application.StatusBar = Format((1 / 4) * 100, "#0") & "%"
    Sheets("Events").Activate
    Cells.Select
    Selection.ClearContents
    Cells(1, 1).Select
    Application.StatusBar = Format((2 / 4) * 100, "#0") & "%"
    Sheets("Cycles").Activate
    Cells.Select
    Selection.ClearContents
    Cells(1, 1).Select
    Application.StatusBar = Format((3 / 4) * 100, "#0") & "%"
    Sheets("Test Info").Activate
    Cells.Select
    Selection.ClearContents
    Cells(1, 1).Select
    Application.StatusBar = Format((4 / 4) * 100, "#0") & "%"
    If SheetExists("WattsTrace") = True Then Sheets("WattsTrace").Delete
    If SheetExists("T1 WattsTrace") = True Then Sheets("T1 WattsTrace").Delete
    If SheetExists("T2 WattsTrace") = True Then Sheets("T2 WattsTrace").Delete
    Sheets("Data Sheet").Activate
    Range("B1:B21").Clear
    Cells(2, 3) = 1
Application.DisplayAlerts = True
Application.StatusBar = False
Application.DisplayStatusBar = oldStatusBar
Application.ScreenUpdating = True
MsgBox ("Clear Complete")
End Sub
Private Sub commandbutton1_Click()
    AHAM
End Sub
Sub GraphWatts()

Dim I As Integer
Dim z As Integer
Dim StartT2Chart As Integer
Dim EndT2Chart As Integer
Dim ChartCount As Integer
Dim RNG1 As Range
Application.DisplayAlerts = False
    'Remove old charts
    'ChartCount = Charts.Count
    'For J = 1 To ChartCount
     '   Charts(J).Delete
    'Next J
    If SheetExists("WattsTrace") = True Then Sheets("WattsTrace").Delete
    If SheetExists("T1 WattsTrace") = True Then Sheets("T1 WattsTrace").Delete
    If SheetExists("T2 WattsTrace") = True Then Sheets("T2 WattsTrace").Delete
    Worksheets("Raw Data").Select
    I = 2
    'Find Last row in Events sheet
    Do While Range("A" & I) <> ""
        I = I + 1
    Loop
    I = I - 1
    
    z = 2
    Do Until Cells(z, 1) >= T1TimeOn Or z >= I
       z = z + 1
    Loop
    StartT1Chart = z
    
    z = 2
    Do Until Cells(z, 1) >= T1TimeOff Or z >= I
       z = z + 1
    Loop
    EndT1Chart = z
 
    
    'Set Chart Range
    Set RNG1 = Range(Cells(StartT1Chart, Wattscol), Cells(EndT1Chart, Wattscol))
    '*****************************
        'Show full Chart
        Dim K As Double
        'Find last row of data
        K = EndT1Chart - StartT1Chart
        'Do Until Sheets("Raw Data").Range("A" & K) = ""
         '   K = K + 1
        'Loop
        'RowNum = K - 1
        
        'Adaptive Scaling
        'K = Abs(DateDiff("n", Sheets("Raw Data").Range("A" & RowNum), Sheets("Raw Data").Range("A" & 2))) 'Data Range
        If K > 30000 Then
            K = 600
        ElseIf K > 20000 Then
            K = 300
        ElseIf K > 10000 Then
            K = 240
        ElseIf K > 5000 Then
            K = 120
        Else:
            K = 60
        End If
        
        Charts.Add
        ActiveChart.ChartType = xlXYScatterLinesNoMarkers
        ActiveChart.SetSourceData Source:=Sheets("Raw Data").Range(RNG1, RNG1), PlotBy:=xlColumns
        ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="T1 WattsTrace"
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
            '.MinorUnit = 10
            .MajorUnit = K
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        
        Selection.TickLabels.Orientation = xlDownward
        Selection.TickLabels.NumberFormat = "General"
        ActiveChart.PlotArea.Height = 470
        
        With Selection.TickLabels.Font
            .Name = "Arial"
            .FontStyle = "Regular"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.ApplyDataLabels Type:=xlDataLabelsShowNone, LegendKey:=False
        ActiveChart.Legend.Select
        Selection.Delete
        ActiveChart.Deselect
        Sheets("Raw Data").Select
        
'*************Create T2 Watts Chart*****************************
    z = 2
    Do Until Cells(z, 1) >= T2TimeOn Or z >= I
       z = z + 1
    Loop
    StartT2Chart = z
    
    z = 2
    Do Until Cells(z, 1) >= T2TimeOff Or z >= I
       z = z + 1
    Loop
    EndT2Chart = z
 

    
    'Set Chart Range
    Set RNG1 = Range(Cells(StartT2Chart, Wattscol), Cells(EndT2Chart, Wattscol))
    '*****************************

        'Find last row of data
        K = EndT2Chart - StartT2Chart
        'Do Until Sheets("Raw Data").Range("A" & K) = ""
         '   K = K + 1
        'Loop
        'RowNum = K - 1
        
        'Adaptive Scaling
        'K = Abs(DateDiff("n", Sheets("Raw Data").Range("A" & RowNum), Sheets("Raw Data").Range("A" & 2))) 'Data Range
        If K > 30000 Then
            K = 600
        ElseIf K > 20000 Then
            K = 300
        ElseIf K > 10000 Then
            K = 240
        ElseIf K > 5000 Then
            K = 120
        Else:
            K = 60
        End If
        
        Charts.Add
        ActiveChart.ChartType = xlXYScatterLinesNoMarkers
        ActiveChart.SetSourceData Source:=Sheets("Raw Data").Range(RNG1, RNG1), PlotBy:=xlColumns
        ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="T2 WattsTrace"
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
            '.MinorUnit = 10
            .MajorUnit = K
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        
        Selection.TickLabels.Orientation = xlDownward
        Selection.TickLabels.NumberFormat = "General"
        ActiveChart.PlotArea.Height = 470
        
        With Selection.TickLabels.Font
            .Name = "Arial"
            .FontStyle = "Regular"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.ApplyDataLabels Type:=xlDataLabelsShowNone, LegendKey:=False
        ActiveChart.Legend.Select
        Selection.Delete
        ActiveChart.Deselect
        Sheets("Raw Data").Select
End Sub

Function SheetExists(SheetName As String) As Boolean
' returns TRUE if the sheet exists in the active workbook
    SheetExists = False
    On Error GoTo NoSuchSheet
    If Len(Sheets(SheetName).Name) > 0 Then
        SheetExists = True
        Exit Function
    End If
NoSuchSheet:

End Function
Function IsOpenWorkbook(strWorkbookName As String, Optional bFullname As Boolean) As Boolean
Dim wb As Workbook
Dim strName As String
    IsOpenWorkbook = False
    For Each wb In Workbooks
        If bFullname = False Then
            strName = wb.Name
        Else
            strName = wb.FullName
        End If
        If (StrComp(strName, strWorkbookName, vbTextCompare) = 0) Then
            IsOpenWorkbook = True
            Exit Function
        End If
    Next
End Function
Private Sub ImportData()

    Dim isel As Integer
    Dim colcount As Integer
    Dim remcntr As Integer
    Dim caption As Variant
    
    
    Dim startModification '(scan for first defrost and discard data)
    Dim proto As Range
    Dim unitOn As Integer
    Dim defrostOn As Integer
    Dim watt As Integer
    Dim T As Integer
    Dim strtmp As String
    Dim strcs As String
    Dim rngSearch As Range
    Dim x As Variant
    
    Dim Time1 As Date
    Dim K As Integer
    Dim L As Integer
    
    Dim A As Variant
     
    Application.ScreenUpdating = False
    
    ' add events sheet
    If Not SheetExists("Import File Info") Then
        Sheets.Add
        ActiveSheet.Name = "Import File Info"
    End If
    If Not SheetExists("Test Info") Then
        Sheets.Add
        ActiveSheet.Name = "Test Info"
    End If
    If Not SheetExists("Events") Then
        Sheets.Add
        ActiveSheet.Name = "Events"
    End If
    If Not SheetExists("Raw Data") Then
        Sheets.Add
        ActiveSheet.Name = "Raw Data"
    End If
    Sheets("Raw Data").Activate
    Dim wkbCalc As Workbook
    Dim wkbImp As Workbook
    Dim RNG1 As Range
    Dim NewFN As Variant
    Dim tsel As Boolean
    Dim tempRowNum As Integer
    Set wkbCalc = ThisWorkbook
    NewFN = Application.GetOpenFilename(FileFilter:="CSV Files (*.csv), *.csv", Title:="Please select a file")
    If NewFN = False Then
        ' Cancel routine
        MsgBox "Macro stopped. CSV File not selected."
        Sheets("Data Sheet").Select
        Exit Sub
    Else
        Workbooks.Open Filename:=NewFN
    End If
    'Pause 5
    Set wkbImp = ActiveWorkbook
    ' extract data from headers and import into a data sheet
    ' isel is row index
    isel = 1
    ' tsel is flag to detect found expression
    tsel = False
    'wkbImp.

    
    Do While tsel = False
        ' get string from the cell
        strtmp = Left(Cells(isel, 1).Value, 6)
        ' compare
        strcs = StrComp(strtmp, "DATA::")
        If strcs = 0 Then
            tsel = True
        Else
            isel = isel + 1
        End If
        If isel > 50 Then
            MsgBox ("data format is invalid, DATA:: not detected or header corrupted")
            Exit Sub
        End If
    Loop
    Range(wkbCalc.Sheets("Raw Data").Cells(1, 1), wkbCalc.Sheets("Raw Data").Cells(25000, 100)).Clear
    Range(wkbCalc.Sheets("Events").Cells(1, 1), wkbCalc.Sheets("Events").Cells(25000, 100)).Clear
    Range(wkbCalc.Sheets("Cycles").Cells(1, 1), wkbCalc.Sheets("Cycles").Cells(25000, 100)).Clear
    Range(wkbCalc.Sheets("Test Info").Cells(1, 1), wkbCalc.Sheets("Test Info").Cells(25000, 100)).Clear
    strtmp = Right(Cells(isel, 1).Value, (Len(Cells(isel, 1)) - 6))
    Cells((isel + strtmp + 1), 1).Select
    Selection.End(xlDown).Select
    ' copy #rows - 1, because last row might be incomplete if data file is downloaded
    ' at the same time as server writes a new scan
    RowNum = ActiveCell.Row - 1
    tempRowNum = RowNum
    ' case last entry is an event
    If Cells(RowNum, 2) = "" Then
        Do While Cells(RowNum, 2) = ""
            RowNum = RowNum - 1
            Cells(RowNum, 2).Select
        Loop
    End If
        ' check for number of columns (channels)
    ' select fixed row (fix from invalid scan or event)
    Cells(RowNum, 1).Select
    Selection.End(xlToRight).Select
    ColNum = ActiveCell.Column
    'start copying data
    ' copy header information (channels)
    Set RNG1 = Range(Cells(1, 1), Cells(isel - 1, ColNum))
    RNG1.Copy Destination:=wkbCalc.Sheets("Test Info").Cells(1, 1)
    
    If Trim(Cells(isel + 1, 1)) = "" Then
        isel = isel + 1
    End If
    
    Set RNG1 = Range(Cells(isel + 1, 1), Cells(isel + 1, ColNum))
    RNG1.Copy Destination:=wkbCalc.Sheets("Raw Data").Cells(1, 1)
    'rng1.Copy Destination:=wkbCalc.Sheets("Events").Cells(1, 1)
    ' copy data for processing

    Set RNG1 = Range(Cells(isel + strtmp + 1, 1), Cells(tempRowNum, ColNum))
    RNG1.Copy Destination:=wkbCalc.Sheets("Raw Data").Cells(2, 1)
    ' clean events from data
    wkbCalc.Activate
    
    'Find new last row
    isel = 2
    Do Until Cells(isel, 1) = ""
        isel = isel + 1
    Loop
    RowNum = isel - 1

    'remove event rows from start of data.  Check 5 rows from start
    isel = 2
    Do While Cells(isel, 2) = "" And isel <= 5
        Rows(isel).Delete
        RowNum = RowNum - 1
    Loop
    '*********************************************************************************
    ' row index
    isel = 1
    ' detection flag
    tsel = False
    ' column detection count (for columns @ -9.90E+37), when treshold reached column gets deleted
    colcount = 1
    'remote counter for events
    remcntr = 1
    ' mark isel as column index
    ' this part removes all columns with open tc's
    ' once it is detected, it loops through the same column number and deletes if tc is open
    ' this behavior is necessary if multiple columns with open tc's are neighbours.
    ' in this case, once column is deleted, other columns shift to left into it's place
    ' we automatically decrease column number in order to save on processing
    For isel = 2 To ColNum
        'Do While (Cells(2, isel).Value = -9.9E+37 Or Cells(4, isel).Value = -9.9E+37 Or Cells(6, isel).Value = -9.9E+37 Or IsEmpty(Cells(1, isel))) And isel < ColNum (changed in v1.9a / updated in v1.9c)
        Do While (IsEmpty(Cells(1, isel)) Or _
        ((Cells(2, isel).Value = -9.9E+37 Or Cells(2, isel).Value = "") And (Cells(4, isel).Value = -9.9E+37 Or Cells(4, isel).Value = "") _
        And (Cells(6, isel).Value = -9.9E+37 Or Cells(6, isel).Value = ""))) _
        And isel < ColNum
            ' delete column
            Columns(isel).EntireColumn.Delete
            ' decrease number of columns
            ColNum = ColNum - 1
        Loop
        '  after:=Cells(2, isel),
        Set rngSearch = Range(Cells(2, isel), Cells(RowNum, isel))
        Set x = rngSearch.Find(what:="-9.9E+37", LookIn:=xlFormulas, LookAt:= _
            xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False)
            If x Is Nothing Then
            
            Else
                ' delete entire row when readings in that tc are detected
                Do While Not x Is Nothing
                    Rows(x.Row).Delete
                    RowNum = RowNum - 1
                    Set x = Nothing
                    Set x = rngSearch.Find(what:="-9.9E+37", LookIn:=xlFormulas, LookAt:= _
                        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                        False)
                Loop
            End If
    Next

    ' this part checks for missing events and inserts them where needed
    ' find defrost, watt unit on columns
    isel = 1
    unitOn = 0
    defrostOn = 0
    watt = 0
    tsel = False
    Do While tsel = False
        If Cells(1, isel).Value = "Event 1 Unit On" Then
            unitOn = isel
        ElseIf Cells(1, isel).Value = "Event 2 Defrost" Then
            defrostOn = isel
        ElseIf Cells(1, isel).Value = "Watt-Hours" Then
            watt = isel
        ElseIf Cells(1, isel).Value = "Watts" Then
            watts = isel
        End If
            
        isel = isel + 1
        
        If isel > ColNum Then
            tsel = True
        End If
    Loop
    
    'Removed the code below on 6/7/13 (KS)
    'Set Defrost On values to "0" when Unit On values are "0" (GG - 2/19/07)
   ' For isel = 2 To RowNum
    '    If Cells(isel, unitOn) = 0 And Cells(isel, defrostOn) = 1 Then
     '       Cells(isel, defrostOn) = 0
      '  End If
    'Next isel
    
    '************************************************************
    ' search for missing UnitOn events by detecting transitions without event lines
    For isel = 2 To RowNum
        ' if there is event transition
        If (Cells(isel, unitOn).Value = 1 And Cells(isel + 1, unitOn).Value = 0) Or (Cells(isel, unitOn).Value = 0 And Cells(isel + 1, unitOn).Value = 1) Then
            ' and there are no events logged after 2 minutes ( data might be buffered, therefore the 2 min span)
            If Not IsEmpty(Cells(isel + 1, 2)) And Not IsEmpty(Cells(isel + 2, 2)) Then
                Rows(isel + 1).Select
                ' insert blank line
                Selection.Insert Shift:=xlDown
                ' take over time and watt-hours from last/next scan depending on transition
                
                If Cells(isel + 2, unitOn).Value = 0 Then
                    ' if there is on/off transition, take readings from first off cycle scan
                    Cells(isel + 1, 1).Value = Cells(isel + 2, 1).Value
                    Cells(isel + 1, watt).Value = Cells(isel + 2, watt).Value
                Else
                    ' else if there is off/on transition, take the reading from last off cycle (neglible non-accumulative error possible)
                    Cells(isel + 1, 1).Value = Cells(isel, 1).Value
                    Cells(isel + 1, watt).Value = Cells(isel, watt).Value
                End If
                ' take over event states and counters from next scan
                Cells(isel + 1, unitOn).Value = Cells(isel + 2, unitOn).Value
                Cells(isel + 1, defrostOn).Value = Cells(isel + 2, defrostOn).Value
                Cells(isel + 1, unitOn + 2).Value = Cells(isel + 2, unitOn + 2).Value
                Cells(isel + 1, unitOn + 3).Value = Cells(isel + 2, unitOn + 3).Value
                Cells(isel + 1, unitOn + 4).Value = Cells(isel + 2, unitOn + 4).Value
                Cells(isel + 1, unitOn + 5).Value = Cells(isel + 2, unitOn + 5).Value
                Cells(isel + 1, unitOn + 6).Value = "Edited"
                isel = isel + 1
            End If
        End If
    Next

    'Check for more anamalous events in Raw Data (v1.14a)
    For I = 2 To RowNum
        If Cells(I, 23) = 1 Then GoTo skipraw
        If Cells(I, 2) = "" And Cells(I + 1, 2) = "" And I < RowNum Then
            If Cells(I, unitOn) <> Cells(I - 1, unitOn) Or Cells(I, unitOn) <> Cells(I + 1, unitOn) And _
            DateDiff("s", Cells(I, 1), Cells(I + 1, 1)) < 60 Then
                'Rows(I).Delete
                'RowNum = RowNum - 1
                'I = I - 1
            End If
        End If
        If ((Cells(I, unitOn) = Cells(I - 1, unitOn) And Cells(I, 2) = "" Or _
         Cells(I, unitOn) <> Cells(I - 1, unitOn) And Cells(I, 2) <> "") And _
         Cells(I, defrostOn) = Cells(I - 1, defrostOn)) And I < RowNum Then
            'Rows(I).Delete
            'RowNum = RowNum - 1
            'I = I - 1
        End If
    skipraw:
    Next I
    
    'extract events into a separate sheet
    isel = 2
    tsel = False
    Do While tsel = False
        If IsEmpty(Cells(isel, 2)) Then
            If IsEmpty(Cells(isel, 1)) Then
                ' end of data reached
                tsel = True
            Else
                ' event detected
                ' increase event counter in "Events" Sheet
                remcntr = remcntr + 1
                Set RNG1 = Range(Cells(isel, 1), Cells(isel, ColNum + 1))
                RNG1.Copy Destination:=wkbCalc.Sheets("Events").Cells(remcntr, 1)
                Rows(isel).Delete Shift:=xlUp
            End If
        Else
            ' increase row index
            isel = isel + 1
        End If
    Loop
    
    
    
    If remcntr = 1 Then
    '    MsgBox ("No events have been detected, variable speed compressor assumed but no defrost detected. Test failed!")
    '    wkbImp.Close
    '    Exit Sub
    '    Changed in v1.9g to handle 100% run
         GoTo Skip100Run
    End If

    ' copy headers
    Set RNG1 = Range(Cells(1, 1), Cells(1, ColNum))
    RNG1.Copy Destination:=wkbCalc.Sheets("Events").Cells(1, 1)
    'clean up event data
    isel = 2
    wkbCalc.Sheets("Events").Activate
    Do While isel < ColNum
        If IsEmpty(Cells(2, isel)) And Not IsEmpty(Cells(1, isel)) Then
            Do While IsEmpty(Cells(2, isel)) And Not IsEmpty(Cells(1, isel))
                Columns(isel).EntireColumn.Delete
            Loop
        End If
        isel = isel + 1
    Loop
    ' redefine isel as event count
    Cells(2, 1).Select
    Selection.End(xlDown).Select
    isel = ActiveCell.Row
    
    
    'If no Defrost event found, prompt, then terminate macro - GG 3/9/07
    If IsEmpty(Cells(2, 1)) Then
        Range("A1").Select
        Sheets("Data Sheet").Activate
        MsgBox ("No Defrost Events found.  Data invalid")
        Exit Sub
    End If

    ' fix def/dwell event handlers
    For I = 2 To (isel - 1)
        If Cells(I, 3).Value = 1 And Cells(I, 4).Value = 1 Then
            If Cells(I + 1, 3).Value = 1 And Cells(I + 1, 4) = 0 Then
                Cells(I + 1, 3).Value = 0
            End If
        End If
    Next I



    ' remove false/duplicate events
    'this section removes the second event occuring when 2 events occur at the same time stamp
    I = 2
    Do Until Cells(I + 1, 1) = ""
        If I < 2 Then I = 2
        If DateDiff("s", Cells(I, 1), Cells(I + 1, 1)) < 120 Then
            If Cells(I + 1, 4).Value <> 1 Then
                If Cells(I + 1, 5).Value <> 1 Then
                    If Cells(I, 3) = Cells(I - 1, 3) Then
                        Rows(I).EntireRow.Delete
                        isel = isel - 1
                        I = I - 1
                    Else
                        Rows(I + 1).EntireRow.Delete
                        isel = isel - 1
                        I = I - 1
                    End If
                End If
            End If
        End If
        I = I + 1
    Loop
    
    
    I = 2
    Do Until Cells(I + 1, 1) = ""
        If Cells(I, 3) = Cells(I + 1, 3) Then
            If Cells(I, 4) = Cells(I + 1, 4) Then
                If Cells(I, 5) = Cells(I + 1, 5) Then
                    Rows(I + 1).EntireRow.Delete
                    isel = isel - 1
                    I = I - 2
                End If
            End If
        End If
        I = I + 1
    Loop



    
    ' detect uneven scans ( more than 2 minutes in between")   (Old method. New method added in v1.9e)
    'For I = 2 To RowNum
    '    If Cells(I + 1, 1) <> "" Then
    '        If DateDiff("n", Cells(I, 1), Cells(I + 1, 1)) > 10 Then
    '            Cells(I, 1).Select
    '            MsgBox ("scans not consistent, bad data, examine selected cell on raw data sheet")
    '            ExitFLG = True
    '            EndFLG = True
    '            wkbImp.Close
    '            Exit Sub
    '        End If
    '    End If
    'Next
    
    Skip100Run:
    wkbCalc.Sheets("Raw Data").Activate
    ' detect uneven scans ( more than 2 minutes in between")   (added in v1.9e)
    For I = 2 To RowNum
        If Cells(I + 1, 1) <> "" Then
            If DateDiff("n", Cells(I, 1), Cells(I + 1, 1)) > 10 Then
                Msg = "Inconsistent scans found in data." & Chr(10) & "Would you like to continue importing data anyway?"
                Ans = MsgBox(Msg, vbYesNo, "Inconsistent Scans")
                Select Case Ans
                    Case vbYes
                        Set tempRng = Range(Cells(I + 1, 1), Cells(RowNum, ColNum))
                        tempRng.Clear
                    Case vbNo
                        Cells(I, 1).Select
                        ExitFLG = True
                        EndFLG = True
                        wkbImp.Close
                        Exit Sub
                End Select
            End If
        End If
    Next
    
    'Call sub to normalize parameter Names (v1.14c)
    CorrectNames
    
    If ExitFLG Then Exit Sub
    
    Set RNG1 = Nothing

    wkbImp.Close savechanges:=False
    Set wkbImp = Nothing
    Set wkbCalc = Nothing
    
    'Call GraphWatts Sub
    'Not calling graphwatts now until summary 2/19/14 KS
    InitialGraphWatts
    '*******************
    Application.ScreenUpdating = True
    AHAM
    
End Sub
Sub CorrectNames()
Dim K As Integer

For I = 2 To ColNum 'Cycle thru all Columns in Raw Data
    For J = 1 To 9 'Cycle thru all possible parameter names in "ParameterNames" sheet
        If Sheets("ParameterNames").Cells(1, J) = Cells(1, I) Then
            GoTo NextParam
        Else:
            K = 1
            Do Until Sheets("ParameterNames").Cells(K, J) = ""
                If Sheets("ParameterNames").Cells(K, J) = Cells(1, I) Then
                    Sheets("Raw Data").Cells(1, I) = Sheets("ParameterNames").Cells(1, J)
                End If
                K = K + 1
            Loop
        End If
    
    Next J
NextParam:
Next I

Columns.AutoFit

End Sub
Function backupimport()
    Dim fso
    Dim f
    Dim FileInfo As String
    Dim proto As Range
    Dim unitOn As Integer
    Dim defrostOn As Integer
    Dim watt As Integer
    Dim T As Integer
    
    
    ' add events sheet
    If Not SheetExists("Import File Info") Then
        Sheets.Add
        ActiveSheet.Name = "Import File Info"
    End If
    If Not SheetExists("Test Info") Then
        Sheets.Add
        ActiveSheet.Name = "Test Info"
    End If
    If Not SheetExists("Events") Then
        Sheets.Add
        ActiveSheet.Name = "Events"
    End If
    If Not SheetExists("Raw Data") Then
        Sheets.Add
        ActiveSheet.Name = "Raw Data"
    End If
    Sheets("Raw Data").Activate
    Dim wkbCalc As Workbook
    Dim wkbImp As Workbook
    Dim RNG1 As Range
    Dim NewFN As Variant
    Dim tsel As Boolean
    Dim tempRowNum As Integer
    Set wkbCalc = ThisWorkbook
    FileInfo = ""
    NewFN = Application.GetOpenFilename(FileFilter:="CSV Files (*.csv), *.csv", Title:="Please select a file")
    If NewFN = False Then
        ' Cancel routine
        MsgBox "Macro stopped. CSV File not selected."
        Exit Sub
    Else
        Workbooks.Open Filename:=NewFN

    End If
    Set wkbImp = ActiveWorkbook
    ' extract data from headers and import into a data sheet
    ' isel is row index
    isel = 1
    ' tsel is flag to detect found expression
    tsel = False
    'wkbImp.
    Do While tsel = False
        ' get string from the cell
        strtmp = Left(Cells(isel, 1).Value, 6)
        ' compare
        strcs = StrComp(strtmp, "DATA::")
        If strcs = 0 Then
            tsel = True
        Else
            isel = isel + 1
        End If
        If isel > 50 Then
            MsgBox ("data format is invalid, DATA:: not detected or header corrupted")
            Exit Sub
        End If
    Loop
    Range(wkbCalc.Sheets("Raw Data").Cells(1, 1), wkbCalc.Sheets("Raw Data").Cells(25000, 100)).Clear
    Range(wkbCalc.Sheets("Events").Cells(1, 1), wkbCalc.Sheets("Events").Cells(25000, 100)).Clear
    Range(wkbCalc.Sheets("Cycles").Cells(1, 1), wkbCalc.Sheets("Cycles").Cells(25000, 100)).Clear
    Range(wkbCalc.Sheets("Test Info").Cells(1, 1), wkbCalc.Sheets("Test Info").Cells(25000, 100)).Clear
    strtmp = Right(Cells(isel, 1).Value, (Len(Cells(isel, 1)) - 6))
    Cells((isel + strtmp + 1), 1).Select
    Selection.End(xlDown).Select
    ' copy #rows - 1, because last row might be incomplete if data file is downloaded
    ' at the same time as server writes a new scan
    RowNum = ActiveCell.Row - 1
    tempRowNum = RowNum
     
    ' case last entry is an event
    If Cells(RowNum, 2) = "" Then
        Do While Cells(RowNum, 2) = ""
            RowNum = RowNum - 1
            Cells(RowNum, 2).Select
        Loop
    End If
        ' check for number of columns (channels)
    ' select fixed row (fix from invalid scan or event)
    Cells(RowNum, 1).Select
    Selection.End(xlToRight).Select
    ColNum = ActiveCell.Column
    'start copying data
    ' copy header information (channels)
    Set RNG1 = Range(Cells(1, 1), Cells(isel - 1, ColNum))
    RNG1.Copy Destination:=wkbCalc.Sheets("Test Info").Cells(1, 1)
    Set RNG1 = Range(Cells(isel + 1, 1), Cells(isel + 1, ColNum))
    RNG1.Copy Destination:=wkbCalc.Sheets("Raw Data").Cells(1, 1)
    'rng1.Copy Destination:=wkbCalc.Sheets("Events").Cells(1, 1)
    ' copy data for processing

    Set RNG1 = Range(Cells(isel + strtmp + 1, 1), Cells(tempRowNum, ColNum))
    RNG1.Copy Destination:=wkbCalc.Sheets("Raw Data").Cells(2, 1)
    ' clean events from data
    wkbCalc.Activate
    ' row index
    isel = 1
    ' detection flag
    tsel = False
    ' column detection count (for columns @ -9.90E+37), when treshold reached column gets deleted
    colcount = 1
    'remote counter for events
    remcntr = 1
    ' mark isel as column index
    ' this part removes all columns with open tc's
    ' once it is detected, it loops through the same column number and deletes if tc is open
    ' this behavior is necessary if multiple columns with open tc's are neighbours.
    ' in this case, once column is deleted, other columns shift to left into it's place
    ' we automatically decrease column number in order to save on processing
    For isel = 2 To ColNum
        Do While (Cells(2, isel).Value = -9.9E+37 Or IsEmpty(Cells(1, isel))) And isel < ColNum
            ' delete column
            Columns(isel).EntireColumn.Delete
            ' decrease number of columns
            ColNum = ColNum - 1
        Loop
    Next
    
    ' this part checks for missing events and inserts them where needed
    ' find defrost, watt unit on columns
    isel = 1
    unitOn = 0
    defrostOn = 0
    watt = 0
    tsel = False
    Do While tsel = False
        If Cells(1, isel).Value = "Event 1 Unit On" Then
            unitOn = isel
        ElseIf Cells(1, isel).Value = "Event 2 Defrost" Then
            defrostOn = isel
        ElseIf Cells(1, isel).Value = "Watt-Hours" Then
            watt = isel
        End If
            
        isel = isel + 1
        
        If isel > ColNum Then
            tsel = True
        End If
    Loop
    ' search for missing events by detecting transitions without event lines
    For isel = 2 To RowNum
        If (Cells(isel, unitOn).Value = 1 And Cells(isel + 1, unitOn).Value = 0) Or (Cells(isel, unitOn).Value = 0 And Cells(isel + 1, unitOn).Value = 1) Then
            If Not IsEmpty(Cells(isel + 1, 2)) Then
                Rows(isel + 1).Select
                Selection.Insert Shift:=xlDown
                ' take over time and watt-hours from last scan
                Cells(isel + 1, 1).Value = Cells(isel, 1).Value
                Cells(isel + 1, watt).Value = Cells(isel, watt).Value
                ' take over event states and counters from next scan
                Cells(isel + 1, unitOn).Value = Cells(isel + 2, unitOn).Value
                Cells(isel + 1, defrostOn).Value = Cells(isel + 2, defrostOn).Value
                Cells(isel + 1, unitOn + 2).Value = Cells(isel + 2, unitOn + 2).Value
                Cells(isel + 1, unitOn + 3).Value = Cells(isel + 2, unitOn + 3).Value
                Cells(isel + 1, unitOn + 4).Value = Cells(isel + 2, unitOn + 4).Value
                Cells(isel + 1, unitOn + 5).Value = Cells(isel + 2, unitOn + 5).Value
                Cells(isel + 1, unitOn + 6).Value = "Edited"
                isel = isel + 1
            End If
        End If
    Next
    
    
    'Exit Sub
    'extract events into a separate sheet
    isel = 2
    tsel = False
    Do While tsel = False
        If IsEmpty(Cells(isel, 2)) Then
            If IsEmpty(Cells(isel, 1)) Then
                ' end of data reached
                tsel = True
            Else
                ' event detected
                ' increase event counter in "Events" Sheet
                remcntr = remcntr + 1
                Set RNG1 = Range(Cells(isel, 1), Cells(isel, ColNum + 1))
                RNG1.Copy Destination:=wkbCalc.Sheets("Events").Cells(remcntr, 1)
                Rows(isel).Delete Shift:=xlUp
            End If
        Else
            ' increase row index
            isel = isel + 1
        End If
    Loop
    If remcntr = 1 Then
        MsgBox ("No events have been detected, variable speed compressor assumed but no defrost detected. Test failed!")
        wkbImp.Close
        Exit Sub
    End If
    ' copy headers
    Set RNG1 = Range(Cells(1, 1), Cells(1, ColNum))
    RNG1.Copy Destination:=wkbCalc.Sheets("Events").Cells(1, 1)
    'clean up event data
    isel = 2
    wkbCalc.Sheets("Events").Activate
    Do While isel < ColNum
        If IsEmpty(Cells(2, isel)) And Not IsEmpty(Cells(1, isel)) Then
            Do While IsEmpty(Cells(2, isel)) And Not IsEmpty(Cells(1, isel))
                Columns(isel).EntireColumn.Delete
            Loop
        End If
        isel = isel + 1
    Loop
    ' redefine isel as event count
    Cells(2, 1).Select
    Selection.End(xlDown).Select
    isel = ActiveCell.Row
    ' remove duplicate events
    For I = 2 To isel
        If (Cells(I, 3).Value = Cells(I - 1, 3).Value And Cells(I, 4).Value = Cells(I - 1, 4).Value And Cells(I, 5).Value = Cells(I - 1, 5).Value) And Not IsEmpty(Cells(I, 1)) Then
            Do While (Cells(I, 3).Value = Cells(I - 1, 3).Value And Cells(I, 4).Value = Cells(I - 1, 4).Value And Cells(I, 5).Value = Cells(I - 1, 5).Value) And Not IsEmpty(Cells(I, 1))
                Rows(I).EntireRow.Delete
                isel = isel - 1
            Loop
        End If
    Next
    ' remove false events
    T = 2
    ' move to first defrost event, earlier data will not be used
    Do While Cells(T, 4) = 0 And Not IsEmpty(Cells(T, 4))
        T = T + 1
    Loop
    If IsEmpty(Cells(T, 4)) Then
        MsgBox ("No defrost event detected, test failed")
        wkbImp.Close
        Exit Sub
    End If
    ' clean up false raising edge event after defrost event if it exists
    If Cells(T + 1, 3).Value = 1 And Cells(T + 1, 4) = 0 Then
        Rows(T + 1).EntireRow.Delete
    End If
    For I = T + 1 To isel
        'if there is a wHrs difference < 2 between events
        
        If (Cells(I, 2) - Cells(I - 1, 2) <= 2) And Not IsEmpty(Cells(I, 1)) Then
            ' event line might be triggered for no reason
            If Cells(I + 1, 2) - Cells(I - 1, 2) <= 2 Then
                ' and less than 2 minutes, => false alarm, delete events
                If DateDiff("n", Cells(I - 1, 1), Cells(I + 1, 1)) <= 2 Then
                    Rows(I + 1).Delete
                    Rows(I).Delete
                    
                End If
            Else
                If DateDiff("n", Cells(I - 1, 1), Cells(I, 1)) <= 2 And Cells(I - 1, 4) <> Cells(I, 4) Then
                    Rows(I).Select
                    MsgBox "ee"
                    Rows(I).Delete
                    Rows(I - 1).Delete
                End If
            End If
        End If
        
        If Not IsEmpty(Cells(I + 1, 1)) Then
            If DateDiff("n", Cells(I, 1), Cells(I + 1, 1)) <= 2 Then
                Rows(I + 1).Delete
                Rows(I).Delete
            End If
        End If
    Next
    wkbCalc.Sheets("Raw Data").Activate
    ' detect uneven scans ( more than 2 minutes in between")
    For I = 2 To RowNum
        If Cells(I + 1, 1) <> "" Then
            If DateDiff("n", Cells(I, 1), Cells(I + 1, 1)) > 10 Then
                Cells(I, 1).Select
                MsgBox ("scans not consistent, bad data, examine selected cell on raw data sheet")
                ExitFLG = True
                wkbImp.Close
                Exit Sub
            End If
        End If
    Next
    
    Set RNG1 = Nothing

    wkbImp.Close
    Set wkbImp = Nothing
    Set wkbCalc = Nothing
    
    'Go to main Sub
    AHAM
End Function

Sub InitialGraphWatts()

Dim I As Integer
Dim ChartCount As Integer
Dim RNG1 As Range
Application.DisplayAlerts = False
    'Remove old charts
    ChartCount = Charts.Count
    For J = 1 To ChartCount
        Charts(J).Delete
    Next J
    Worksheets("Raw Data").Select
    I = 2
    'Find Last row in Events sheet
    Do While Range("A" & I) <> ""
        I = I + 1
    Loop
    I = I - 1
    'Set Chart Range
    Set RNG1 = Range(Cells(2, watts), Cells(I, watts))
    '*****************************
        'Show full Chart
        Dim K As Double
        'Find last row of data
        K = 1
        Do Until Sheets("Raw Data").Range("A" & K) = ""
            K = K + 1
        Loop
        RowNum = K - 1
        
        'Adaptive Scaling
        K = Abs(DateDiff("n", Sheets("Raw Data").Range("A" & RowNum), Sheets("Raw Data").Range("A" & 2))) 'Data Range
        If K > 30000 Then
            K = 600
        ElseIf K > 20000 Then
            K = 300
        ElseIf K > 10000 Then
            K = 240
        ElseIf K > 5000 Then
            K = 120
        Else:
            K = 60
        End If
        
        Charts.Add
        ActiveChart.ChartType = xlXYScatterLinesNoMarkers
        ActiveChart.SetSourceData Source:=Sheets("Raw Data").Range(RNG1, RNG1), PlotBy:=xlColumns
        ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="WattsTrace"
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
            '.MinorUnit = 10
            .MajorUnit = K
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        
        Selection.TickLabels.Orientation = xlDownward
        Selection.TickLabels.NumberFormat = "General"
        ActiveChart.PlotArea.Height = 470
        
        With Selection.TickLabels.Font
            .Name = "Arial"
            .FontStyle = "Regular"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.ApplyDataLabels Type:=xlDataLabelsShowNone, LegendKey:=False
        ActiveChart.Legend.Select
        Selection.Delete
        ActiveChart.Deselect
        Sheets("Raw Data").Select
End Sub


Sub ChooseCycleAmount()

On Error GoTo errhandle


Msg = "Would you like to select which cycles to use for T1?"
Ans = MsgBox(Msg, vbYesNo, "Modify number of cycles")
    
Select Case Ans
    Case vbYes
        ChooseCycleBool = True
        GoTo Continue
    Case vbNo
        ChooseCycleBool = False
        Exit Sub
End Select

    
Continue:
    CycleAmount = InputBox("Please enter the start cycle you would like to use:", "Start of T1")

    If CycleAmount > FullCycle - 1 Or CycleAmount < J Then
        Ans = MsgBox("Please choose a value between " & J & " - " & FullCycle - 1, vbOKOnly, "Invalid Input")
        GoTo Continue
    End If
    
Continue2:
    CycleAmount2 = InputBox("Please enter the end cycle you would like to use:", "End of T1")

    If CycleAmount2 > FullCycle Or CycleAmount2 <= CycleAmount Then
        Ans = MsgBox("Please choose a value between " & CycleAmount + 1 & " - " & FullCycle, vbOKOnly, "Invalid Input")
        GoTo Continue2
    End If

If ChooseCycleBool = False Then
errhandle:
    ChooseCycleBool = False
    Exit Sub
End If

FFBW1Avg = 0
FFBW2Avg = 0
FFBW3Avg = 0
FZBW1Avg = 0
FZBW2Avg = 0
FZBW3Avg = 0
FZBW4Avg = 0
FZBW5Avg = 0

  For I = CycleAmount To CycleAmount2
                FFBW1Avg = FFBW1Avg + FFBW1(I)
                FFBW2Avg = FFBW2Avg + FFBW2(I)
                FFBW3Avg = FFBW3Avg + FFBW3(I)
                FZBW1Avg = FZBW1Avg + FZBW1(I)
                FZBW2Avg = FZBW2Avg + FZBW2(I)
                FZBW3Avg = FZBW3Avg + FZBW3(I)
                FZBW4Avg = FZBW4Avg + FZBW4(I)
                FZBW5Avg = FZBW5Avg + FZBW5(I)
            Next I
            FFBW1Avg = FFBW1Avg / (I - CycleAmount)
            FFBW2Avg = FFBW2Avg / (I - CycleAmount)
            FFBW3Avg = FFBW3Avg / (I - CycleAmount)
            FFAvg = (FFBW1Avg + FFBW2Avg + FFBW3Avg) / 3
            FZBW1Avg = FZBW1Avg / (I - CycleAmount)
            FZBW2Avg = FZBW2Avg / (I - CycleAmount)
            FZBW3Avg = FZBW3Avg / (I - CycleAmount)
            If FZBW4ch <> "NA" Then
                FZBW4Avg = FZBW4Avg / (I - CycleAmount)
            End If
            If FZBW5ch <> "NA" Then
                FZBW5Avg = FZBW5Avg / (I - CycleAmount)
            End If
            
            If FZBW4ch <> "NA" And FZBW5ch <> "NA" Then
                FZAvg = (FZBW1Avg + FZBW2Avg + FZBW3Avg + FZBW4Avg + FZBW5Avg) / 5
            ElseIf FZBW4ch <> "NA" Then
                FZAvg = (FZBW1Avg + FZBW2Avg + FZBW3Avg + FZBW4Avg) / 4
            ElseIf FZBW5ch <> "NA" Then
                FZAvg = (FZBW1Avg + FZBW2Avg + FZBW3Avg + FZBW5Avg) / 4
            Else
                FZAvg = (FZBW1Avg + FZBW2Avg + FZBW3Avg) / 3
            End If
            
            EP1 = (WattHoursCount(CycleAmount2 + 1) - WattHoursCount(CycleAmount)) / 1000
            t1 = 24 * 60 * (TimeCycleOn(CycleAmount2 + 1) - TimeCycleOn(CycleAmount))
            StartT1 = CycleAmount
            T1TimeOn = TimeCycleOn(StartT1)
            EndT1 = CycleAmount2
            T1TimeOff = TimeCycleOn(EndT1 + 1)
            
            EP2calc
            ECalc
            

End Sub


Sub Run100()

Dim stab_time As Double
Dim start_A As Integer
Dim end_A As Integer
Dim start_gap As Integer
Dim end_gap As Integer
Dim start_B As Integer
Dim end_B As Integer
Dim stab_bool As Boolean
Dim stab_temp As Double
Dim stabA_avg As Double
Dim stabB_avg As Double
Dim rng As Range
Dim cnt As Integer
Sheets("Raw Data").Activate
AvailChannel
    
start_A = 1
Do While Not stab_bool:
    start_A = start_A + 1
    end_A = start_A
    Do While Abs(DateDiff("s", Cells(start_A, 1).Value, Cells(end_A, 1).Value)) < 7200
        end_A = end_A + 1
    Loop
    start_gap = end_A
    end_gap = end_A
    Do While Abs(DateDiff("s", Cells(start_gap, 1).Value, Cells(end_gap, 1).Value)) < 10800
        end_gap = end_gap + 1
    Loop
    start_B = end_gap
    end_B = end_gap
    Do While Abs(DateDiff("s", Cells(start_B, 1).Value, Cells(end_B, 1).Value)) < 7200
        end_B = end_B + 1
        If end_B > RowNum Then
            MsgBox ("Stability not found.")
            End
        End If
    Loop
    
    stab_time = (Abs(DateDiff("s", Cells(start_B, 1).Value, Cells(end_B, 1).Value)) / 2) + Abs(DateDiff("s", Cells(start_gap, 1).Value, Cells(end_gap, 1).Value)) _
    + (Abs(DateDiff("s", Cells(start_A, 1).Value, Cells(end_A, 1).Value)) / 2)

    stab_bool = True
    stab_temp = 0.042 * (stab_time / 60 / 60)
      
    For I = LBound(channelarray) To UBound(channelarray)
        If InStr(channelarray(I), "BW") <> 0 Then
            Set rng = Range(Cells(start_A, I), Cells(end_A, I))
            stabA_avg = Application.Average(rng)
            Set rng = Range(Cells(start_B, I), Cells(end_B, I))
            stabB_avg = Application.Average(rng)
            If Not Abs(stabA_avg - stabB_avg) <= stab_temp Then
                stab_bool = False
            End If
        End If
    Next I
Loop
    
If Not stab_bool Then
    MsgBox ("Stability not found.")
    End
End If
    
ReDim IndexCycleOn(1)
ReDim TimeCycleOn(1)
ReDim IndexCycleOff(1)
ReDim TimeCycleOff(1)
ReDim WattHoursCount(1)

I = start_A
cnt = 1

Do While I < RowNum
    Do While Abs(DateDiff("s", Cells(I, 1).Value, Cells(I + cnt, 1).Value)) < 1799
        If UBound(IndexCycleOn) < 2 Then
            IndexCycleOn(1) = I
            TimeCycleOn(1) = Cells(I, 1).Value
            Exit Do
        End If
        cnt = cnt + 1
        If I + cnt > RowNum Then Exit Do
    Loop
    If I + cnt > RowNum Then Exit Do
    If UBound(IndexCycleOn) > 1 Then
        I = I + cnt
        cnt = 1
        IndexCycleOn(UBound(IndexCycleOn)) = I
        TimeCycleOn(UBound(TimeCycleOn)) = Cells(I, 1).Value
    End If
    ReDim Preserve IndexCycleOn(UBound(IndexCycleOn) + 1)
    ReDim Preserve TimeCycleOn(UBound(TimeCycleOn) + 1)
    Do While Abs(DateDiff("s", Cells(I, 1).Value, Cells(I + cnt, 1).Value)) < 1799
        cnt = cnt + 1
        If I + cnt > RowNum Then Exit Do
    Loop
    If I + cnt > RowNum Then Exit Do
    I = I + cnt
    cnt = 1
    IndexCycleOff(UBound(IndexCycleOff)) = I
    TimeCycleOff(UBound(TimeCycleOff)) = Cells(I, 1).Value
    WattHoursCount(UBound(WattHoursCount)) = Cells(I - 1, WattHourscol).Value
    ReDim Preserve IndexCycleOff(UBound(IndexCycleOff) + 1)
    ReDim Preserve TimeCycleOff(UBound(TimeCycleOff) + 1)
    ReDim Preserve WattHoursCount(UBound(WattHoursCount) + 1)
Loop

If IndexCycleOn(UBound(IndexCycleOn)) = 0 Then
    ReDim Preserve IndexCycleOn(UBound(IndexCycleOn) - 1)
    ReDim Preserve TimeCycleOn(UBound(TimeCycleOn) - 1)
End If
If IndexCycleOff(UBound(IndexCycleOff)) = 0 Then
    ReDim Preserve IndexCycleOff(UBound(IndexCycleOff) - 1)
    ReDim Preserve TimeCycleOff(UBound(TimeCycleOff) - 1)
End If
If WattHoursCount(UBound(WattHoursCount)) = 0 Then
    ReDim Preserve WattHoursCount(UBound(WattHoursCount) - 1)
End If

FullCycle = UBound(IndexCycleOn) - 1
J = UBound(IndexCycleOn)


End Sub














