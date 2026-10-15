Attribute VB_Name = "Setup"
' functions in this module are
' GetCycles()
' SetupVars()
' AvailChannel()
' RowCol()
'Option Explicit

Public PreDwell As Boolean
Public TimePreDwellOn As Date
Public TimePreDwellOff As Date
Public MLgradLeft As Double
Public MLgradLeftMax As Double
Public MLgradRight As Double
Public MLgradRightMax As Double
Public HMgradLeft As Double
Public HMgradLeftMax As Double
Public HMgradRight As Double
Public HMgradRightMax As Double

Sub GetCycles(startCY As Integer)
    Dim rngSearch As Range
    Dim x As Variant
    Dim searchFlg As Boolean
    Dim TimeDef2On As Variant
    Dim def2OnFlg As Boolean
    Dim eventcount As Integer
    Dim wattsct As Integer
    Dim avgcount As Integer
    Dim FFerr As Boolean
    Dim TCycleOn As Date
    Dim TCycleOff As Date
        
    
    searchFlg = True
    def2OnFlg = False
    EventTimeStampCol = 1
    EventWattHourscol = 2
    eventunitoncol = 3
    EventDefrostoncol = 4
    Sheets("Events").Activate
    Cells(1.1).Select
    Selection.End(xlDown).Select
    If ActiveCell.Row < 99999 Then
        eventcount = ActiveCell.Row
    End If
    TimeDef2On = 0
    cyclic = True
    wattsct = 0
    
    LastScanIndex = RowNum
    LastScan = Sheets("Raw Data").Cells(RowNum, 1)
    
    If AdjustedStart = False Then
        Start = 2
    Else
        z = 2
        Do
            z = z + 1
        Loop Until Cells(z, 1) >= TimeToStart Or z >= RowNum
        Start = z
    End If
    

    '*************************************************************************************************************************************
    If eventcount < 1 Then
    run100bool = True
    Run100  'Need a go to line
    GoTo Run100NextLine
    End If
    ' looking only for cycle start (on / off times)
    Set rngSearch = Range(Cells(Start, eventunitoncol), Cells(eventcount, eventunitoncol))
    J = 1
    ReDim TimeCycleOn(J)
    ReDim TimeCycleOff(J)
    ReDim WattHoursCount(J)
    
    ' for all data on a sheet
    
    Do While Not IsEmpty(Cells(Start, 1))
        ' reset counter for cycles
        TCycleOn = 0
        TCycleOff = 0
        wattsct = 0
    ' find cycle on event
        Set x = Nothing
        Set x = rngSearch.Find(what:="1", After:=Cells(Start, eventunitoncol), LookIn:=xlValues, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False)
        If x Is Nothing Then
            MsgBox ("No Cycle On Event detected, Test Failed")
            ExitFLG = True
            Exit Sub
        Else
            If Start > x.Row Then
                
                Exit Do
            End If
            
            If DateDiff("n", Cells(x.Row, 1), Sheets("Raw Data").Cells(LastScanIndex, 1)) <= 0 Then
                
                Exit Do
            End If
            Start = x.Row
            TCycleOn = Cells(Start, 1)
            wattsct = Cells(Start, 2)
        End If
        Set x = Nothing
        Set x = rngSearch.Find(what:="0", After:=Cells(Start, eventunitoncol), LookIn:=xlValues, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False)
        If x Is Nothing Then
            MsgBox ("Data error, no cycle off detected, Test failed")
            ExitFLG = True
            Exit Sub
        Else
            ' exit loop if index is lower, end of detection
            If Start > x.Row Then
                If J = 1 Then
                    cyclic = False
                    TimeCycleOn(1) = TCycleOn
                    

                    TimeCycleOff(1) = Cells(RowNum, 1)
                Else
                    ' on time where last cycle is not completed
                    LastScan = TCycleOn
                End If
                
                Exit Do
            End If
            If DateDiff("s", Cells(x.Row, 1), Sheets("Raw Data").Cells(LastScanIndex, 1)) / 60 <= 0 Then
                If J = 1 Then
                    cyclic = False
                    TimeCycleOn(1) = TCycleOn
                    TimeCycleOff(1) = Cells(LastScanIndex, 1)
                Else
                    LastScan = TCycleOn
                End If
                Exit Do
            End If
            Start = x.Row
            TCycleOff = Cells(Start, 1)
        End If
        ' if there is no on cycle, discard last cycle and mark lastscan time
        Set x = Nothing
        Set x = rngSearch.Find(what:="1", After:=Cells(Start, eventunitoncol), LookIn:=xlValues, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False)
        If x.Row < Start Then
            LastScan = TCycleOn
            Exit Do
        End If
        If DateDiff("s", Cells(x.Row, 1), Sheets("Raw Data").Cells(LastScanIndex, 1)) / 60 <= 0 Then
            LastScan = TCycleOn
            Exit Do
        End If
        
        ReDim Preserve TimeCycleOn(J)
        ReDim Preserve TimeCycleOff(J)
        ReDim Preserve WattHoursCount(J)
        TimeCycleOn(J) = TCycleOn
        WattHoursCount(J) = wattsct
        TimeCycleOff(J) = TCycleOff
        J = J + 1
    Loop
    ' find index when times start or end
    ' reduce index J to show number of cycles
    J = J - 1


    
    
    Sheets("Raw Data").Activate
    If cyclic = False Then
        Start = 2
        ' FIND START
        Do While DateDiff("s", TimeCycleOn(1), Cells(Start, 1).Value) / 60 <= 0
            Start = Start + 1
        Loop

        If Cells(Start - 1, UnitOncol) = 1 Then
            Start = Start - 1
        End If
        
        J = (LastScanIndex - Start) / 60

        
        tsel = False
        isel = 1
        Do While tsel = False
            If Cells(1, isel).Value = "Watt-Hours" Then
                watt = isel
            End If
                
            isel = isel + 1
            
            If isel > ColNum Then
                tsel = True
            End If
        Loop
        
        
        ReDim IndexCycleOn(J)
        ReDim TimeCycleOn(J)
        ReDim WattHoursCount(J)
        
        

        For I = 1 To J
            If I > 1 Then
                Dt = 0
                IndexCycleOn(I) = IndexCycleOn(I - 1) + 60
                If DateDiff("n", Cells(IndexCycleOn(I - 1), 1), Cells(IndexCycleOn(I), 1)) > 60 Then
                    Do While DateDiff("n", Cells(IndexCycleOn(I - 1), 1), Cells(IndexCycleOn(I) - Dt, 1)) > 60
                        Dt = Dt + 1
                    Loop
                    IndexCycleOn(I) = IndexCycleOn(I) - Dt
                End If
            Else
                IndexCycleOn(I) = Start
            End If
            TimeCycleOn(I) = Cells(IndexCycleOn(I), 1)
            WattHoursCount(I) = Cells(IndexCycleOn(I), watt)
           
        Next
    Else
        Start = 2
        ReDim IndexCycleOn(J)
        ReDim IndexCycleOff(J)

        For I = 1 To J
            Do While DateDiff("s", TimeCycleOn(I), Cells(Start, 1).Value) / 60 <= 0
                Start = Start + 1
            Loop
            If Cells(Start - 1, UnitOncol) = 1 Then
                IndexCycleOn(I) = (Start - 1)
            Else
                IndexCycleOn(I) = Start
            End If
            Do While DateDiff("s", TimeCycleOff(I), Cells(Start, 1).Value) / 60 <= 0
                Start = Start + 1
            Loop
            If Cells(Start - 1, UnitOncol) = 0 Then
                IndexCycleOff(I) = (Start - 1)
            Else
                IndexCycleOff(I) = Start
            End If
        Next






        
        Do While DateDiff("s", LastScan, Cells(Start, 1).Value) / 60 <= 0 And Not IsEmpty(Cells(Start, 1))
            Start = Start + 1
        Loop
        If Cells(Start - 1, 1) = 1 Then
            LastScanIndex = Start - 1
        Else
            LastScanIndex = Start
        End If

    End If
    
    DefOnWattHourCount = 1
    LastWattHourCount = Cells(LastScanIndex, WattHourscol).Value
    Start = 2
    Do While DateDiff("s", TimeDefOn, Cells(Start, 1).Value) / 60 <= 0
        Start = Start + 1
    Loop
    If Cells(Start - 1, 1) = 1 Then
        Start = Start - 1
    End If
  '*****************************************************************************************************************************************************
  
    Run100NextLine:
    ReDim FFBW1(J)
    ReDim FFBW2(J)
    ReDim FFBW3(J)
    ReDim FZBW1(J)
    ReDim FZBW2(J)
    ReDim FZBW3(J)
    ReDim FZBW4(J)
    ReDim FZBW5(J)
    ReDim ICEBW(J)
    ReDim VoltsAv(J)
    ReDim VoltsMin(J)
    ReDim VoltsMax(J)
    ReDim RGrad(J)
    ReDim LGrad(J)
    ReDim RSOCAvg(J)
    ReDim LSOCAvg(J)
    ReDim LSOCHIavg(J)
    ReDim LSOCLOavg(J)
    ReDim RSOCHIavg(J)
    ReDim RSOCLOavg(J)
    'Dim rng As Range
    MLgradLeftMax = 0
    HMgradLeftMax = 0
    MLgradRightMax = 0
    HMgradRightMax = 0

        
    ' get averages
    For I = 1 To J - 1

        Set rng = Range(Cells(IndexCycleOn(I), FFBW1col), Cells(IndexCycleOn(I + 1) - 1, FFBW1col))
        FFBW1(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Set rng = Range(Cells(IndexCycleOn(I), FFBW2col), Cells(IndexCycleOn(I + 1) - 1, FFBW2col))
        FFBW2(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Set rng = Range(Cells(IndexCycleOn(I), FFBW3col), Cells(IndexCycleOn(I + 1) - 1, FFBW3col))
        FFBW3(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Set rng = Range(Cells(IndexCycleOn(I), FZBW1col), Cells(IndexCycleOn(I + 1) - 1, FZBW1col))
        FZBW1(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Set rng = Range(Cells(IndexCycleOn(I), FZBW2col), Cells(IndexCycleOn(I + 1) - 1, FZBW2col))
        FZBW2(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Set rng = Range(Cells(IndexCycleOn(I), FZBW3col), Cells(IndexCycleOn(I + 1) - 1, FZBW3col))
        FZBW3(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        If FZBW4ch <> "NA" Then
            Set rng = Range(Cells(IndexCycleOn(I), FZBW4col), Cells(IndexCycleOn(I + 1) - 1, FZBW4col))
            FZBW4(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Else
            FZBW4(I) = 0
        End If
        If FZBW5ch <> "NA" Then
            Set rng = Range(Cells(IndexCycleOn(I), FZBW5col), Cells(IndexCycleOn(I + 1) - 1, FZBW5col))
            FZBW5(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Else
            FZBW5(I) = 0
        End If
        
        If ICEBWch <> "NA" Then
            Set rng = Range(Cells(IndexCycleOn(I), ICEBWcol), Cells(IndexCycleOn(I + 1) - 1, ICEBWcol))
            ICEBW(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Else
            ICEBW(I) = 0
        End If
        
        If Voltsch <> "NA" Then
            Set rng = Range(Cells(IndexCycleOn(I), Voltscol), Cells(IndexCycleOn(I + 1) - 1, Voltscol))
            VoltsAv(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
            VoltsMin(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Min(rng), 2)
            VoltsMax(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Max(rng), 2)
        Else
            VoltsAv(I) = 0
            VoltsMin(I) = 0
            VoltsMax(I) = 0
        End If
        
        Set rng = Range(Cells(IndexCycleOn(I), RSOCcol), Cells(IndexCycleOn(I + 1) - 1, RSOCcol))
        RSOCAvg(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Set rng = Range(Cells(IndexCycleOn(I), LSOCcol), Cells(IndexCycleOn(I + 1) - 1, LSOCcol))
        LSOCAvg(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Set rng = Range(Cells(IndexCycleOn(I), RSOCHIcol), Cells(IndexCycleOn(I + 1) - 1, RSOCHIcol))
        RSOCHIavg(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Set rng = Range(Cells(IndexCycleOn(I), LSOCHIcol), Cells(IndexCycleOn(I + 1) - 1, LSOCHIcol))
        LSOCHIavg(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Set rng = Range(Cells(IndexCycleOn(I), RSOCLOcol), Cells(IndexCycleOn(I + 1) - 1, RSOCLOcol))
        RSOCLOavg(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        Set rng = Range(Cells(IndexCycleOn(I), LSOCLOcol), Cells(IndexCycleOn(I + 1) - 1, LSOCLOcol))
        LSOCLOavg(I) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        LGrad(I) = Abs(Application.WorksheetFunction.Round((LSOCHIavg(I) - LSOCLOavg(I)) / (hght / 12), 3))
        RGrad(I) = Abs(Application.WorksheetFunction.Round((RSOCHIavg(I) - RSOCLOavg(I)) / (hght / 12), 3))
        MLgradLeft = Abs(Application.WorksheetFunction.Round((LSOCAvg(I) - LSOCLOavg(I)) / (hght / 12), 3))
        MLgradRight = Abs(Application.WorksheetFunction.Round((RSOCAvg(I) - RSOCLOavg(I)) / (hght / 12), 3))
        HMgradLeft = Abs(Application.WorksheetFunction.Round((LSOCHIavg(I) - LSOCAvg(I)) / (hght / 12), 3))
        HMgradRight = Abs(Application.WorksheetFunction.Round((RSOCHIavg(I) - RSOCAvg(I)) / (hght / 12), 3))
        If MLgradLeft > MLgradLeftMax Then
            MLgradLeftMax = MLgradLeft
        End If
        If MLgradRight > MLgradRightMax Then
            MLgradRightMax = MLgradRight
        End If
        If HMgradLeft > HMgradLeftMax Then
            HMgradLeftMax = HMgradLeft
        End If
        If HMgradRight > HMgradRightMax Then
            HMgradRightMax = HMgradRight
        End If
        
    Next
    
    Set rng = Range(Cells(IndexCycleOn(J), FFBW1col), Cells(LastScanIndex, FFBW1col))
    FFBW1(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(IndexCycleOn(J), FFBW2col), Cells(LastScanIndex, FFBW2col))
    FFBW2(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(IndexCycleOn(J), FFBW3col), Cells(LastScanIndex, FFBW3col))
    FFBW3(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(IndexCycleOn(J), FZBW1col), Cells(LastScanIndex, FZBW1col))
    FZBW1(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(IndexCycleOn(J), FZBW2col), Cells(LastScanIndex, FZBW2col))
    FZBW2(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(IndexCycleOn(J), FZBW3col), Cells(LastScanIndex, FZBW3col))
    FZBW3(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    If FZBW4ch <> "NA" Then
        Set rng = Range(Cells(IndexCycleOn(J), FZBW4col), Cells(LastScanIndex, FZBW4col))
        FZBW4(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Else
        FZBW4(J) = 0
    End If
    If FZBW5ch <> "NA" Then
        Set rng = Range(Cells(IndexCycleOn(J), FZBW5col), Cells(LastScanIndex, FZBW5col))
        FZBW5(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Else
        FZBW5(J) = 0
    End If
    
    If ICEBWch <> "NA" Then
        Set rng = Range(Cells(IndexCycleOn(J), ICEBWcol), Cells(LastScanIndex, ICEBWcol))
        ICEBW(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Else
        ICEBW(J) = 0
    End If
    
    If Voltsch <> "NA" Then
        Set rng = Range(Cells(IndexCycleOn(J), Voltscol), Cells(LastScanIndex, Voltscol))
        VoltsAv(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        VoltsMin(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Min(rng), 2)
        VoltsMax(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Max(rng), 2)
    Else
        VoltsAv(J) = 0
        VoltsMin(J) = 0
        VoltsMax(J) = 0
    End If
        
    Set rng = Range(Cells(IndexCycleOn(J), RSOCcol), Cells(LastScanIndex, RSOCcol))
    RSOCAvg(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(IndexCycleOn(J), LSOCcol), Cells(LastScanIndex, LSOCcol))
    LSOCAvg(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    
    Set rng = Range(Cells(IndexCycleOn(J), RSOCHIcol), Cells(LastScanIndex, RSOCHIcol))
    RSOCHIavg(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(IndexCycleOn(J), LSOCHIcol), Cells(LastScanIndex, LSOCHIcol))
    LSOCHIavg(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(IndexCycleOn(J), RSOCLOcol), Cells(LastScanIndex, RSOCLOcol))
    RSOCLOavg(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(IndexCycleOn(J), LSOCLOcol), Cells(LastScanIndex, LSOCLOcol))
    LSOCLOavg(J) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    LGrad(J) = Abs(Application.WorksheetFunction.Round((LSOCHIavg(J) - LSOCLOavg(J)) / (hght / 12), 3))
    RGrad(J) = Abs(Application.WorksheetFunction.Round((RSOCHIavg(J) - RSOCLOavg(J)) / (hght / 12), 3))
    
    ' store number of cycles
    cycleCount = J

    '    TimeDwellOff = Sheets("Raw Data").Cells(IndexCycleOn(1) - 1, 1)

    

        
    ' reset start pointer to point to start of defrost event
    Start = 2
    Do While DateDiff("s", TimeDefOn, Cells(Start, 1).Value) / 60 <= 0
        Start = Start + 1
    Loop
    
    If Not SheetExists("Cycles") Then
        Sheets.Add
        ActiveSheet.Name = "Cycles"
    Else
        Sheets("Cycles").Activate
        Range(Cells(1, 1), Cells(1000, 10)).Clear
    End If
    
    Cells(1, 1).Value = "Cycle"
    Cells(1, 2).Value = "Time On"
    Cells(1, 3).Value = "Index On"
    Cells(1, 4).Value = "Time Off"
    Cells(1, 5).Value = "Index Off"
    Cells(1, 6).Value = "Duty Cycle"
    

    For I = 1 To J
        Cells(I + 1, 1).Value = I
        Cells(I + 1, 2).Value = TimeCycleOn(I)
        Cells(I + 1, 3).Value = IndexCycleOn(I)
        Cells(I + 1, 4).Value = TimeCycleOff(I)
        Cells(I + 1, 5).Value = IndexCycleOff(I)
        If I < J Then
            Cells(I + 1, 6).Value = ((DateDiff("n", TimeCycleOn(I), TimeCycleOff(I))) / (DateDiff("n", TimeCycleOn(I), TimeCycleOn(I + 1)) - 1))
        Else
            Cells(I + 1, 6).Value = ((DateDiff("n", TimeCycleOn(I), TimeCycleOff(I))) / (DateDiff("n", TimeCycleOn(I), LastScan) - 1))
        End If
    Next
    Cells(J + 2, 1).Value = "ls"
    Cells(I + 1, 2).Value = LastScan
    Cells(I + 1, 3).Value = LastScanIndex
    


End Sub
Sub SetupVars()
    Wattsch = Application.WorksheetFunction.VLookup("Watts", Range("tcs"), 2, False)
    WattHoursch = Application.WorksheetFunction.VLookup("Watt-Hours", Range("tcs"), 2, False)
    UnitOnch = Application.WorksheetFunction.VLookup("Unit On", Range("tcs"), 2, False)
    DefrostOnch = Application.WorksheetFunction.VLookup("Defrost On", Range("tcs"), 2, False)
    RSOCHIch = Application.WorksheetFunction.VLookup("RSOC HI", Range("tcs"), 2, False)
    RSOCLOch = Application.WorksheetFunction.VLookup("RSOC LO", Range("tcs"), 2, False)
    RSOCch = Application.WorksheetFunction.VLookup("RSOC Mid", Range("tcs"), 2, False)
    LSOCch = Application.WorksheetFunction.VLookup("LSOC Mid", Range("tcs"), 2, False)
    LSOCHIch = Application.WorksheetFunction.VLookup("LSOC HI", Range("tcs"), 2, False)
    LSOCLOch = Application.WorksheetFunction.VLookup("LSOC LO", Range("tcs"), 2, False)
    FFBW1ch = Application.WorksheetFunction.VLookup("FF BW 1", Range("tcs"), 2, False)
    FFBW2ch = Application.WorksheetFunction.VLookup("FF BW 2", Range("tcs"), 2, False)
    FFBW3ch = Application.WorksheetFunction.VLookup("FF BW 3", Range("tcs"), 2, False)
    FZBW1ch = Application.WorksheetFunction.VLookup("FZ BW 1", Range("tcs"), 2, False)
    FZBW2ch = Application.WorksheetFunction.VLookup("FZ BW 2", Range("tcs"), 2, False)
    FZBW3ch = Application.WorksheetFunction.VLookup("FZ BW 3", Range("tcs"), 2, False)
    FZBW4ch = Application.WorksheetFunction.VLookup("FZ BW 4", Range("tcs"), 2, False)
    FZBW5ch = Application.WorksheetFunction.VLookup("FZ BW 5", Range("tcs"), 2, False)
    ICEBWch = Application.WorksheetFunction.VLookup("ICE BW", Range("tcs"), 2, False)
    Voltsch = Application.WorksheetFunction.VLookup("Volts", Range("tcs"), 2, False)

    
    For J = 1 To chan
        If channelarray(J) = Wattsch Then Wattscol = channel(J, 2)
        If channelarray(J) = WattHoursch Then WattHourscol = channel(J, 2)
        If channelarray(J) = UnitOnch Then UnitOncol = channel(J, 2)
        If channelarray(J) = DefrostOnch Then DefrostOncol = channel(J, 2)
        If channelarray(J) = RSOCHIch Then RSOCHIcol = channel(J, 2)
        If channelarray(J) = RSOCLOch Then RSOCLOcol = channel(J, 2)
        If channelarray(J) = RSOCch Then RSOCcol = channel(J, 2)
        If channelarray(J) = LSOCch Then LSOCcol = channel(J, 2)
        If channelarray(J) = LSOCHIch Then LSOCHIcol = channel(J, 2)
        If channelarray(J) = LSOCLOch Then LSOCLOcol = channel(J, 2)
        If channelarray(J) = FFBW1ch Then FFBW1col = channel(J, 2)
        If channelarray(J) = FFBW2ch Then FFBW2col = channel(J, 2)
        If channelarray(J) = FFBW3ch Then FFBW3col = channel(J, 2)
        If channelarray(J) = FZBW1ch Then FZBW1col = channel(J, 2)
        If channelarray(J) = FZBW2ch Then FZBW2col = channel(J, 2)
        If channelarray(J) = FZBW3ch Then FZBW3col = channel(J, 2)
        If channelarray(J) = FZBW4ch Then FZBW4col = channel(J, 2)
        If channelarray(J) = FZBW5ch Then FZBW5col = channel(J, 2)
        If channelarray(J) = ICEBWch Then ICEBWcol = channel(J, 2)
        If channelarray(J) = Voltsch Then Voltscol = channel(J, 2)
    Next J
End Sub

Sub AvailChannel()
'Get All Available Channel Names
    Row = 1: Col = 1: chan = 0
    ReDim channelarray(1 To ColNum) As String
    ' ct is value at row 3 column 7
    CT = Range("CT").Value
    For I = 1 To ColNum
        Name = Cells(Row, I)
        If Not Name = "" Then
            chan = chan + 1
            channel(chan, 1) = Val(Name)
            channelarray(chan) = Name
            channel(chan, 2) = I
        End If
    Next I

End Sub

Sub RowCol()
'Get Number of Rows and Columns of Data from Data Cart
    Sheets(RawDataName).Select
    Selection.End(xlDown).Select
    RowNum = ActiveCell.Row
    Selection.End(xlToRight).Select
    ColNum = ActiveCell.Column
    ReDim channel(ColNum + 1, 2)
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Sheets("Events").Select
    Cells(1, 1).Select
    Selection.End(xlDown).Select
    If ActiveCell.Row < 99999 Then
        EventsCount = ActiveCell.Row
    End If
    Sheets("Raw Data").Select
    If RowNum < 600 Then
        Msg = "insufficient scans, increase duration (600 scans limit), current is " & RowNum
        MsgBox (Msg)
        ExitFLG = True
    End If
End Sub
