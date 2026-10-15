Attribute VB_Name = "AHAMCalc"
Option Explicit
Option Base 1

Public AmbientFLG As Boolean
Public Ans As Variant
Public avgrow As Integer
Public AvgWatts As Double
Public C As Integer
Dim caption As String
Public chan As Integer
Public channel() As Integer
Public channelarray() As String
Public ChooseCycleBool As Boolean
Public Col As Integer
Public ColNum As Integer
Public ComputerName As String
Public CreateFlag As Boolean
Public CT As Double
Public CycleAmount As Long
Public CycleAmount2 As Long
Public cycleCount As Integer
Public cycles() As Integer
Public cyclic As Boolean
Public DebugMode As Boolean
Public DefrostCol As Integer
Public DefFFBW1 As Double
Public DefFFBW2 As Double
Public DefFFBW3 As Double
Public DefFZBW1 As Double
Public DefFZBW2 As Double
Public DefFZBW3 As Double
Public DefFZBW4 As Double
Public DefFZBW5 As Double
Public DefICEBW As Double
Public DefLGrad As Double
Public DefLSOC As Double
Public DefRGrad As Double
Public FFDefrostCol As Integer
Public DefrostOnch As Variant
Public DefrostOncol As Integer
Public DefRSOC As Double
Public DefVoltsAv As Double
Public DefVoltsMax As Double
Public DefVoltsMin As Double
Public Delta As Integer
Public Difference As Double
Public DwellFFBW1 As Double
Public DwellFFBW2 As Double
Public DwellFFBW3 As Double
Public DwellFZBW1 As Double
Public DwellFZBW2 As Double
Public DwellFZBW3 As Double
Public DwellFZBW4 As Double
Public DwellFZBW5 As Double
Public DwellICEBW As Double
Public DwellLGrad As Double
Public DwellLSOC As Double
Public DwellRGrad As Double
Public DwellRSOC As Double
Public DwellVoltsAv As Double
Public DwellVoltsMax As Double
Public DwellVoltsMin As Double
Public EndDef As Integer
Public EndDwell As Integer
Public EndFLG As Boolean
Public EndRow As Integer
Public EndT1 As Integer
Public EndT2 As Integer
Public EndT3 As Integer
Public EP1 As Double
Public EP1Watts As Range
Public EP2 As Double
Public EP2Watts As Range
Public ET As Double
Public ETcorr As Double
Public EventsCount As Integer
Public EventCountCol As Integer
Public ExitFLG As Boolean
Public FFAvg As Double
Public FFBW1() As Double
Public FFBW11 As Range
Public FFBW12avg  As Double
Public FFBW1ch  As String
Public FFBW1col  As Integer
Public FFBW2() As Double
Public FFBW21 As Range
Public FFBW22avg  As Double
Public FFBW2ch As String
Public FFBW2col As Integer
Public FFBW3() As Double
Public FFBW31 As Range
Public FFBW32avg  As Double
Public FFBW3ch As String
Public FFBW3col As Integer
Public FFCT As Double
Public FZAvg As Double
Public FZBW1() As Double
Public FZBW11 As Range
Public FZBW11avg  As Double
Public FZBW12avg  As Double
Public FZBW1ch As Variant
Public FZBW1col As Integer
Public FZBW2() As Double
Public FZBW21 As Range
Public FZBW21avg  As Double
Public FZBW22avg  As Double
Public FZBW2ch As Variant
Public FZBW2col As Integer
Public FZBW3() As Double
Public FZBW31 As Range
Public FZBW31avg  As Double
Public FZBW32avg  As Double
Public FZBW3ch As Variant
Public FZBW3col As Integer
Public FZBW4() As Double
Public FZBW41 As Range
Public FZBW41avg  As Double
Public FZBW42avg  As Double
Public FZBW4ch As Variant
Public FZBW4col As Variant
Public FZBW5() As Double
Public FZBW51 As Range
Public FZBW51avg  As Double
Public FZBW52avg  As Double
Public FZBW5ch As Variant
Public FZBW5col As Variant
Public GradientFLG As Boolean
Public GradientLimit As Double
Public hght As Double
Public I As Integer
Public ICEBW() As Double
Public ICEBWch As Variant
Public ICEBWcol As Variant
Public IndexCycleOff() As Integer
Public IndexCycleOn() As Integer
Public IndexDwellOff As Integer
Public J As Integer
Public LastEvent As Integer
Public LastRow As Integer
Public LastScan As Date
Public LastScanIndex As Integer
Public LastWattHourCount As Integer
Public LGrad() As Double
Public LSGrad As Double
Public LSOCAvg() As Double
Public LSOCch As Variant
Public LSOCcol As Integer
Public LSOCHIavg() As Double
Public LSOCHIcol As Integer
Public LSOCLOavg() As Double
Public LSOCLOcol As Integer
Public LSOCMax As Single
Public LSOCMin As Single
Public LSOCTotAvg As Double
Public Msg As String
Public Name As String
Public NextDefrost As Integer
Public NumCycles As Integer
Public offset As Integer
Public OffTime(1 To 500) As Double
Public OnCol As Integer
Public OnTime(1 To 500) As Double
Public R As Integer
Public RawDataName As String
Public ReportName As String
Public RGrad() As Double
Public Row As Integer
Public RowNum As Integer
Public RSGrad As Double
Public RSOCAvg() As Double
Public RSOCch As Variant
Public RSOCcol As Integer
Public RSOCHIavg() As Double
Public RSOCHIcol As Integer
Public RSOCLOavg() As Double
Public RSOCLOcol As Integer
Public RSOCMax As Single
Public RSOCMin As Single
Public RSOCTotAvg As Double
Public run100bool As Boolean
Public stabcycle As Integer
Public Stable As Boolean
Public Start As Integer
Public StartDef As Integer
Public StartDwell As Integer
Public StartT1 As Integer
Public StartT2 As Integer
Public StartT3 As Integer
Public StepOne As Boolean
Public StepReport As Boolean
Public StepThree As Boolean
Public StepTwo As Boolean
Public SumWatts As Double
Public t1 As Double
Public T1TimeOn As Date
Public T1TimeOff As Date
Public StartT1Chart As Integer
Public EndT1Chart As Integer
Public T2 As Double
Public T2DefOff As Integer
Public T2DefOn As Integer
Public t2defwatton As Double
Public T2FF1(1 To 500) As Double
Public T2FF2(1 To 500) As Double
Public T2FF3(1 To 500) As Double
Public T2FF4(1 To 500) As Double
Public T2FF5(1 To 500) As Double
Public T2FZ1(1 To 500) As Double
Public T2FZ2(1 To 500) As Double
Public T2FZ3(1 To 500) As Double
Public T2FZ4(1 To 500) As Double
Public T2FZ5(1 To 500) As Double
Public T2IB(1 To 500) As Double
Public T2LGrad(1 To 500) As Double
Public T2LsocAvg(1 To 500) As Double
Public T2Off(1 To 500) As Date
Public T2OffTime(1 To 500) As Double
Public T2On(1 To 500) As Date
Public T2OnTime(1 To 500) As Double
Public T2Post As Integer
Public T2PostEnd(0 To 500) As Integer
Public T2PostStart(0 To 500) As Integer
Public T2Pre As Integer
Public T2PreCycleEnd(0 To 500) As Integer
Public T2PreCycleStart(0 To 500) As Integer
Public T2Rgrad(1 To 500) As Double
Public T2RsocAvg(0 To 500) As Double
Public T2TimeOff As Date
Public T2TimeOn As Date
Public T2Volts(0 To 500) As Double
Public T2VoltsMax(1 To 500) As Double
Public T2VoltsMin(0 To 500) As Double
Public T2WH(1 To 500) As Double
Public T2WHOff As Double
Public T2WHOn As Double
Public TempEnd(0 To 500) As Integer
Public TempStart(0 To 500) As Integer
Public tempRng As Range
Public TestDuration As Double
Public TimeDefOff As Date
Public TimeDefOn As Date
Public TimeDwellOff As Date
Public TimeDwellOn As Date
Public TimeCycleOff() As Date
Public TimeCycleOn() As Date
Public TimeStart() As Date
Public UnitOnch As Variant
Public UnitOncol As Integer
Public UserName As String
Public VoltsAv() As Double
Public Voltsch As Variant
Public Voltscol As Variant
Public VoltsMax() As Double
Public VoltsMin() As Double
Public WattHoursCount() As Integer
Public WattRange As Range
Public Wattsch As Variant
Public Wattscol As Integer
Public WattsDefOff As Double
Public WattsDefOn As Double
Public WattHoursch As Variant
Public WattHourscol As Integer
Public WHCol As Integer
Public wkbReport As Workbook
Public wkbThis As Workbook
Public FZOnly As Boolean
Public FZOnlyKfactor As Double
Public FZOnlyPrompt As Boolean
Public TimeToStart As Date
Public AdjustedStart As Boolean



Sub AHAM()
Attribute AHAM.VB_ProcData.VB_Invoke_Func = " \n14"
    
    ChooseCycleBool = False
    FZOnlyPrompt = True
    wbPath = ActiveWorkbook.Path
    Set wkbThis = ThisWorkbook
    wkbThis.Sheets("Raw Data").Activate
    Range("A1").Select
    RawDataName = ActiveSheet.Name
    If Range("A1").Value <> "Timestamp" Then
        Sheets("Data Sheet").Select
        Msg = "This program must be run from "
        Msg = Msg & "the worksheet that has the "
        Msg = Msg & "raw data from the data cart."
        caption = "ERROR!"
        If MsgBox(Msg, vbOKOnly, caption) = vbOK Then End
    End If
    ' get gradient parametars
    DebugMode = True
    Sheets("Data Sheet").Select
    hght = Cells(10, 7).Value 'height is a technician input
    hght = hght + 10  '10" adjustment for t/c offset
    GradientFLG = True
    AmbientFLG = True
    GradientLimit = 0.5 * hght / 12
    ExitFLG = False
    EndFLG = False
    CreateFlag = True
    On Error GoTo CreateSheets
    FFCT = Cells(8, 7).Value
    caption = "Create New Data Sheet?"
    Msg = "Would you like to use the values "
    Msg = Msg & "in your current Data Sheet?"
    
    Ans = MsgBox(Msg, vbYesNo, caption)
    
    Select Case Ans
        Case vbYes
            CreateFlag = False
            GoTo CreateSheets
        Case vbNo
            CreateFlag = True
            GoTo CreateSheets
    End Select
    

    
CreateSheets:
    On Error GoTo 0
    RowCol
    If ExitFLG = True Then Exit Sub
    AvailChannel
    If CreateFlag = True Then
        AddDataSheet
        Prompt
        End
    End If
    CheckUserInput
    If ExitFLG = True Then Exit Sub
        
    Msg = "Would you like to start at the beginning of the data?"
    Ans = MsgBox(Msg, vbYesNo, caption)
    
    Select Case Ans
        Case vbYes
            AdjustedStart = False
            GoTo ContinueTest
        Case vbNo
            AdjustedStart = True
            TimeToStart = InputBox("Enter the date & time to start (e.g., 2/7/14 4:15 am)")
            GoTo ContinueTest
    End Select
        

ContinueTest:
        
        
    SetupVars
    FZOnlyKfactor = Sheets("Data Sheet").Cells(2, 3)
    ' so we can move back 2 positions from last defrost event
    NextDefrost = 1
    Do While EndFLG = False
        GetCycles (NextDefrost)
        If ExitFLG = True Then Exit Sub
        If FZOnly = False Then
            'Removed in v1.9b
            'Sheets("Data Sheet").Cells(2, 3).Value = 1
            'FZOnlyKfactor = Sheets("Data Sheet").Cells(2, 3)
            StabilityCheck
            If Stable = False Then
                MsgBox ("Test not stable")
                Sheets("Data Sheet").Activate
                End
            End If
        Else:
            StabilityCheckFZOnly
        End If
        If ExitFLG = True Then Exit Sub
    Loop
    If Stable = False Then
        MsgBox ("test not stable")
    End If
    Application.ScreenUpdating = True
      
End Sub
Sub CheckUserInput()
    Sheets("Data Sheet").Activate
    For I = 1 To 21
        If Cells(I, 2).Value = "" Then
            Cells(I, 2).Select
            MsgBox ("Please select proper channel or NA if channel is not available")
            ExitFLG = True
            Exit Sub
        End If
    Next I
    'If Cells(8, 7).Value = 0 Then     'Removed FF portion in V1.6
        'Cells(8, 7).Select
        'MsgBox ("Please enter FFCT value")
        'ExitFLG = True
    'End If
    If Cells(3, 7).Value = 0 Then
            Cells(3, 7).Select
            MsgBox ("Please enter CT value")
            ExitFLG = True
    End If
End Sub

Sub AddDataSheet()
Attribute AddDataSheet.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim flgFC As Boolean            ' flag found channel
    flgFC = False
    On Error Resume Next
    Application.DisplayAlerts = False
      
    Application.DisplayAlerts = True
    On Error GoTo 0
  
    Sheets("Data Sheet").Select
    Range("B2:B21,C2,G1:G3").Clear
    
    Col = 100
    For R = 1 To chan
        Cells(R, Col) = channelarray(R)
    Next R
    Cells(R, Col) = "NA"
    
    Range("A1:C1,C2").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    Range("C2").Select
    Selection.Font.ColorIndex = 3
    Selection.Interior.ColorIndex = 15
    
    Range("a1").Value = "TC or Power"
    Range("b1").Value = "Channel Name"
    Range("C1").Value = "K Factor(FZ Only)"
    Range("C2").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="1,.85,.7"
        End With
    Range("C2").Value = 1
    
    Range("a2").Value = "FF BW 1"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "FF BW 1" Then
            Cells(2, 2) = "FF BW 1"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b2").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & chan
        End With
    End If
    
    Range("a3").Value = "FF BW 2"
        flgFC = False
    For R = 1 To chan
        If channelarray(R) = "FF BW 2" Then
            Cells(3, 2) = "FF BW 2"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b3").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & chan
        End With
    End If
    Range("a4").Value = "FF BW 3"
        flgFC = False
    For R = 1 To chan
        If channelarray(R) = "FF BW 3" Then
            Cells(4, 2) = "FF BW 3"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b4").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & chan
        End With
    End If



    Range("a5").Value = "FZ BW 1"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "FZ BW 1" Then
            Cells(5, 2) = "FZ BW 1"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b5").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & chan
        End With
    End If
    Range("a6").Value = "FZ BW 2"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "FZ BW 2" Then
            Cells(6, 2) = "FZ BW 2"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b6").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & chan
        End With
    End If
    Range("a7").Value = "FZ BW 3"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "FZ BW 3" Then
            Cells(7, 2) = "FZ BW 3"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b7").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & chan
        End With
    End If
    Range("a8").Value = "FZ BW 4"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "FZ BW 4" Then
            Cells(8, 2) = "FZ BW 4"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b8").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b8").Value = "NA"
    End If
    
    Range("a9").Value = "FZ BW 5"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "FZ BW 5" Then
            Cells(9, 2) = "FZ BW 5"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b9").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & ((chan + 1))
        End With
        Range("b9").Value = "NA"
    End If
    
    Range("a10").Value = "ICE BW"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "ICE BW" Then
            Cells(10, 2) = "ICE BW"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b10").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b10").Value = "NA"
    End If

    
    Range("a11").Value = "RSOC LO"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "RSOC LO" Then
            Cells(11, 2) = "RSOC LO"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b11").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b11").Value = "NA"
    End If
    
    Range("a12").Value = "RSOC Mid"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "RSOC MD" Then
            Cells(12, 2) = "RSOC MD"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b12").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b12").Value = "NA"
    End If
    
    Range("a13").Value = "RSOC HI"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "RSOC UP" Then
            Cells(13, 2) = "RSOC UP"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b13").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b13").Value = "NA"
    End If
    
    Range("a14").Value = "LSOC LO"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "LSOC LO" Then
            Cells(14, 2) = "LSOC LO"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b14").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b14").Value = "NA"
    End If
    
    Range("a15").Value = "LSOC Mid"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "LSOC MD" Then
            Cells(15, 2) = "LSOC MD"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b15").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b15").Value = "NA"
    End If
    
    Range("a16").Value = "LSOC HI"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "LSOC UP" Then
            Cells(16, 2) = "LSOC UP"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b16").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b16").Value = "NA"
    End If
    
    Range("a17").Value = "Volts"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "Volts" Then
            Cells(17, 2) = "Volts"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
    Range("b17").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b17").Value = "NA"
    End If
    
    Range("a18").Value = "Watts"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "Watts" Then
            Cells(18, 2) = "Watts"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
    Range("b18").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b18").Value = "NA"
    End If
    
    Range("a19").Value = "Watt-Hours"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "Watt-Hours" Then
            Cells(19, 2) = "Watt-Hours"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b19").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b19").Value = "NA"
    End If
    
    Range("a20").Value = "Unit On"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "Event 1 Unit On" Then
            Cells(20, 2) = "Event 1 Unit On"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b20").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b20").Value = "NA"
    End If
    
    Range("a21").Value = "Defrost On"
    flgFC = False
    For R = 1 To chan
        If channelarray(R) = "Event 2 Defrost" Then
            Cells(21, 2) = "Event 2 Defrost"
            R = chan
            flgFC = True
        End If
    Next R
    If Not flgFC Then
        Range("b21").Select
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$CV$1:$CV$" & (chan + 1)
        End With
        Range("b21").Value = "NA"
    End If
    'Columns("a:c").EntireColumn.AutoFit
    Columns("B").AutoFit
    
    'Range("c4").Value = "Set ALL 3 FF parameters to NA for FZ Only Units"     'Removed FF section in V1.6
    'Range("c4").Select
        'Selection.Font.ColorIndex = 3
        'Selection.Font.Bold = True
        'Selection.Font.Underline = xlUnderlineStyleSingle
    
    Range("A2:A" & ColNum).Select
    ActiveWorkbook.Names.Add Name:="tcs", RefersToR1C1:= _
        "='Data Sheet'!R2C1:R21C2"
        
    Range("F1:F3").Select
        With Selection
            .HorizontalAlignment = xlRight
        End With
    Range("F1").Value = "Min Time Between Defrosts "
    Range("F2").Value = "Max Time Between Defrosts "
    Range("F3").Value = "CT "
    Range("G1").Select
    Range("G1").AddComment
    Range("G1").Comment.Visible = False
    Range("G1").Comment.Text Text:="Input Must be between 6 and 12"
    ActiveWorkbook.Names.Add Name:="CTL", RefersToR1C1:="='Data Sheet'!R1C7"
    With Selection.Validation
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=CTL>=6"
    End With
    Range("G2").Select
    Range("G2").AddComment
    Range("G2").Comment.Visible = False
    Range("G2").Comment.Text Text:="Input Must be between 12 and 96"
    ActiveWorkbook.Names.Add Name:="CTM", RefersToR1C1:="='Data Sheet'!R2C7"
    With Selection.Validation
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=CTM<=96"
    End With
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(CTL=0,CTM=0),0,ROUND(CTL*CTM/(0.2*(CTM-CTL)+CTL),1))"
    ActiveWorkbook.Names.Add Name:="CT", RefersToR1C1:="='Data Sheet'!R3C7"
    Range("B21").Select
    


Range("D15").Select

End Sub
Sub Prompt()
Attribute Prompt.VB_ProcData.VB_Invoke_Func = " \n14"
    Msg = "Please choose the channel number from the pull down list "
    Msg = Msg & Chr(13) & "that corresponds to the TC or Power reading."
    Msg = Msg & Chr(13) & Chr(13)
    Msg = Msg & "Also input the correct values for maximum and minimum compressor,"
    Msg = Msg & Chr(13) & " for the CT calculation."
    Msg = Msg & Chr(13) & Chr(13) & "Rerun the AHAM macro by using the Run AHAM Test button"
    MsgBox Msg
    Range("D22").Select
    
    
End Sub

Sub PartOne()

    StepOne = True
    StepTwo = False
    StepThree = False
    StepReport = False
    AHAM

End Sub

Sub PartTwo()
    'Removed in V1.6
    StepOne = False
    StepTwo = True
    StepThree = False
    StepReport = False
    AHAM
    
End Sub

Sub PartThree()
    'Removed in V1.6
    StepOne = False
    StepTwo = False
    StepThree = True
    StepReport = False
    AHAM
    
End Sub

Sub FinalReport()

    StepOne = False
    StepTwo = False
    StepThree = False
    StepReport = True
    AHAM
    
End Sub
