Attribute VB_Name = "DefrostPortion"
Public FF1Col As Integer
Public FF3Col As Integer
Public FZ1Col As Integer
Public FZ3Col As Integer
Public FZ4Col As Integer
Public FZ5Col As Integer
Public FullCycle As Integer
Public ffStabAvg As Double
Public StabAvg As Double
Public T2AStart As Integer
Public T2AEnd As Integer
Public T2DefEnd As Integer
Public T2DefStart As Integer
Public T2PCEnd As Integer
Public T2PCStart As Integer
Public T2RecEnd As Integer
Public T2RecStart As Integer
Public T2BEnd As Integer
Public T2BStart As Integer
Public T2BPlus As Date
Public t2runtime(1 To 500) As Date
Public T3AStart As Integer
Public T3AEnd As Integer
Public T3DefEnd As Integer
Public T3DefStart As Integer
Public T3PCEnd As Integer
Public T3PCStart As Integer
Public T3RecEnd As Integer
Public T3RecStart As Integer
Public T3BEnd As Integer
Public T3BStart As Integer
Public T3BPlus As Date
Public T3runtime(1 To 500) As Date
Public T2Run(1 To 500) As Double
Public T2RunPer(1 To 500) As Double
Public T3Run(1 To 500) As Double
Public T3RunPer(1 To 500) As Double
Public CycleA As Range
Public CycleAAvg As Double
Public CycleAEnd As Integer
Public CycleAStart As Integer
Public CycleB As Range
Public CycleBAvg As Double
Public CycleBEnd As Integer
Public CycleBStart As Integer
Public DefrostStart As Integer
Public ffcycleA As Range
Public ffcycleB As Range
Public ffcycleBAvg As Double
Public ffcycleAAvg As Double
Public Flag As Integer
Public x As Integer
Public y As Integer


Sub EP2calc()
    
    If run100bool Then Exit Sub
    Sheets("Raw Data").Activate
    Cells(1, 1).Select
    Selection.End(xlDown).Select
    LastRow = ActiveCell.Row
    StabAvg = FZAvg
          
     
    'Setup column variables
    y = 0
    Do
        y = y + 1
    Loop Until Cells(1, y) = "Event 1 Unit On"
    OnCol = y
         
    y = 0
    Do
        y = y + 1
    Loop Until Cells(1, y) = "Event 2 Defrost"
    DefrostCol = y
    
    y = 0
    Do
        y = y + 1
    Loop Until Cells(1, y) = "Event 3 Comp" 'Will need to change Event 3 Comp to actual name once this is determined
    FFDefrostCol = y
    
     y = 0
    Do
        y = y + 1
    Loop Until Cells(1, y) = "Event Count 1 Unit On"
    EventCountCol = y
        
    tempcol = Sheets("Data Sheet").Cells(5, 2)
    y = 0
    Do
        y = y + 1
    Loop Until Cells(1, y) = tempcol
    FZ1Col = y
    
    tempcol = Sheets("Data Sheet").Cells(7, 2)
    y = 0
    Do
        y = y + 1
    Loop Until Cells(1, y) = tempcol
    FZ3Col = y
    
    tempcol = Sheets("Data Sheet").Cells(8, 2)
    y = 0
If tempcol <> "NA" Then
    Do
        y = y + 1
    Loop Until Cells(1, y) = tempcol
    FZ4Col = y
End If
    
    tempcol = Sheets("Data Sheet").Cells(9, 2)
    y = 0
If tempcol <> "NA" Then
    Do
        y = y + 1
    Loop Until Cells(1, y) = tempcol
    FZ5Col = y
End If

    tempcol = Sheets("Data Sheet").Cells(2, 2)
    y = 0
    Do
        y = y + 1
    Loop Until Cells(1, y) = tempcol
    FF1Col = y

    tempcol = Sheets("Data Sheet").Cells(4, 2)
    y = 0
    Do
        y = y + 1
    Loop Until Cells(1, y) = tempcol
    FF3Col = y
    
    tempcol = Sheets("Data Sheet").Cells(19, 2)
    y = 0
    Do
        y = y + 1
    Loop Until Cells(1, y) = tempcol
    WHCol = y
    
    'Find the Freezer defrost
    y = 0
    Do
        y = y + 1
        If y >= LastRow Then
            MsgBox ("No Freezer Defrost in data set.  Test failed.")
            Sheets("Data Sheet").Activate
            End
        End If
    Loop Until (Cells(y, DefrostCol) = 1)
    DefrostStart = y
    T2DefOn = y
    T2DefOff = y
    
    x = DefrostStart - 1

If cyclic = True Then
    Do While Cells(x, OnCol) = 0
        x = x - 1
    Loop
    T2PCEnd = x
    T2DefOff = x
    
    'at this point x=the end of the precool cycle, assuming an offtime after the precool, if no offtime then this loop will be skipped
    Do While Cells(x, OnCol) = 1
        x = x - 1
    Loop
    T2PCStart = x
    'at this point x=the beginning of the precool cycle
    Do While Cells(x, OnCol) = 0
        x = x - 1
    Loop
    'at this point x = the beginning of the off cylce before the precool, this is a potential beginning point of T2
    T2Pre = 1
    CycleAEnd = x ' sets the end point of cycle a
    T2PreCycleEnd(T2Pre) = x
    y = x
    
   'the next two loops set the parameters of cycle A
FindCycleA: 'Will continue looping back through the test to find the required cycle A that is within .5 of the stable temp
    Do While Cells(y, OnCol) = 1
        y = y - 1
        If y < 2 Then
            MsgBox ("Part 2 Cycle A (Before Defrost) never stabilizes. Test Failed")
            Sheets("Data Sheet").Activate
            End
        End If
    Loop
    t2runtime(T2Pre) = Cells(y, 1)
    Do While Cells(y, OnCol) = 0
        y = y - 1
        If y < 2 Then
            MsgBox ("Part 2 Cycle A (Before Defrost) never stabilizes. Test Failed")
            Sheets("Data Sheet").Activate
            End
        End If
    Loop
    CycleAStart = y + 1 'the start of the potential cycle a
    T2PreCycleStart(T2Pre) = y + 1
    
    Set CycleA = Range(Cells(CycleAStart, FZ1Col), Cells(CycleAEnd, FZ3Col))
    
    If FZBW4ch <> "NA" Then
        Set CycleA = Range(Cells(CycleAStart, FZ1Col), Cells(CycleAEnd, FZ4Col))
    End If
    If FZBW5ch <> "NA" Then
        Set CycleA = Range(Cells(CycleAStart, FZ1Col), Cells(CycleAEnd, FZ5Col))
    End If
    
    CycleAAvg = Application.WorksheetFunction.Average(CycleA)
    Set ffcycleA = Range(Cells(CycleAStart, FF1Col), Cells(CycleAEnd, FF3Col))
    ffcycleAAvg = Application.WorksheetFunction.Average(ffcycleA)
    

    If CycleAAvg >= StabAvg - 0.5 And CycleAAvg <= StabAvg + 0.5 And ffcycleAAvg >= FFAvg - 0.5 And ffcycleAAvg <= FFAvg + 0.5 Then
                StartT2 = CycleAEnd + 1 'this will be the start of the t2 period
    Else
        CycleAEnd = CycleAStart - 1
        T2Pre = T2Pre + 1
        T2PreCycleEnd(T2Pre) = CycleAEnd
        GoTo FindCycleA
    End If
    T2AStart = CycleAStart
    T2AEnd = CycleAEnd
    'FullCycle = (Cells(CycleAStart, EventCountCol) - Cells(IndexCycleOn(1), EventCountCol)) / 2
    FullCycle = (Cells(DefrostStart, EventCountCol) - Cells(IndexCycleOn(1), EventCountCol)) / 2
     
'This section finds the end of the calculation period
    x = DefrostStart
    T2DefStart = DefrostStart

    Do While Cells(x, OnCol) = 1 Or Cells(x, DefrostCol) = 1
        x = x + 1
    Loop
    T2DefEnd = x
    'at this point x = the end of the defrost portion
    Do While Cells(x, OnCol) = 0
        x = x + 1
    Loop
    T2RecStart = x
    
    
    'at this point x = the beginning of the recovery cycle
    Do While Cells(x, OnCol) = 1
        x = x + 1
    Loop
    T2Post = 1
    T2RecEnd = x
    T2PostStart(T2Post) = x
    'at this point x = the end of the recovery cycle
     Do While Cells(x, OnCol) = 0
        x = x + 1
    Loop
    'at this point x = the beginning of the first regular cycle after defrost
    t2runtime(4 + T2Pre + T2Post) = Cells(x, 1).Value
    Do While Cells(x, OnCol) = 1
        x = x + 1
    Loop
    T2PostEnd(T2Post) = x - 1
    CycleBStart = x
    y = x
   
   
FindCycleB: 'Will continue looping forward to find the required cycle B that is within .5 of the stable temp
    T2Post = T2Post + 1
    T2PostStart(T2Post) = y
    Do While Cells(y, OnCol) = 0
        y = y + 1
        If y >= LastRow Then
            MsgBox ("Part 2 Cycle B (After Defrost) never stabilizes. Test Failed")
            Sheets("Data Sheet").Activate
            End
        End If
    Loop
    t2runtime(5) = Cells(x, 1).Value
    Do While Cells(y, OnCol) = 1
        y = y + 1
        If y >= LastRow Then
            MsgBox ("Part 2 Cycle B (After Defrost) never stabilizes. Test Failed")
            Sheets("Data Sheet").Activate
            End
        End If
    Loop
    CycleBEnd = y - 1 'the end of the potential cycle b
    T2PostEnd(T2Post) = y - 1
    
    Set CycleB = Range(Cells(CycleBStart, FZ1Col), Cells(CycleBEnd, FZ3Col))
    
    If FZBW4ch <> "NA" Then
        Set CycleB = Range(Cells(CycleBStart, FZ1Col), Cells(CycleBEnd, FZ4Col))
    End If
    If FZBW5ch <> "NA" Then
        Set CycleB = Range(Cells(CycleBStart, FZ1Col), Cells(CycleBEnd, FZ5Col))
    End If
    
    CycleBAvg = Application.WorksheetFunction.Average(CycleB)
    Set ffcycleB = Range(Cells(CycleBStart, FF1Col), Cells(CycleBEnd, FF3Col))
    ffcycleBAvg = Application.WorksheetFunction.Average(ffcycleB)


    
    If CycleBAvg >= StabAvg - 0.5 And CycleBAvg <= StabAvg + 0.5 And ffcycleBAvg >= FFAvg - 0.5 And ffcycleBAvg <= FFAvg + 0.5 Then
            EndT2 = CycleBStart - 1 'this will be the start of the t2 period
        
    Else
        CycleBStart = CycleBEnd + 1
        t2runtime(5 + T2Pre + T2Post) = t2runtime(5)
        GoTo FindCycleB
    End If
    
    T2BStart = CycleBStart
    T2BEnd = CycleBEnd

Else  'noncyclic
    'Find start of T2
    x = DefrostStart
    
    Do
        x = x - 1
         If x < 2 Then
            MsgBox ("Part 2 not stable before defrost. Test Failed.")
            Sheets("Data Sheet").Activate
            End
        End If
    Loop Until Cells(x, Wattscol) >= AvgWatts - 5 And Cells(x, Wattscol) <= AvgWatts + 5
      
    Flag = 0
    Do While Flag = 0
        Set CycleA = Range(Cells(x, FZ1Col), Cells(x, FZ3Col))
                
        If FZBW4ch <> "NA" Then
            Set CycleA = Range(Cells(x, FZ1Col), Cells(x, FZ4Col))
        End If
        If FZBW5ch <> "NA" Then
            Set CycleA = Range(Cells(x, FZ1Col), Cells(x, FZ5Col))
        End If
        
        CycleAAvg = Application.WorksheetFunction.Average(CycleA)
        If CycleAAvg >= StabAvg - 0.5 And CycleAAvg <= StabAvg + 0.5 Then
            StartT2 = x 'this will be the start of the t2 period
            Flag = 1
        Else
            x = x - 1
        End If
    Loop
    
    'Find end of T2
    Flag = 0
    x = DefrostStart
    Do
        x = x + 1
         If x >= LastRow Then
            MsgBox ("Part 2 not stable after defrost.  Test failed.")
            Sheets("Data Sheet").Activate
            End
        End If
    Loop Until Cells(x, Wattscol) >= AvgWatts - 5 And Cells(x, Wattscol) <= AvgWatts + 5
    Do While Flag = 0
        Set CycleB = Range(Cells(x, FZ1Col), Cells(x, FZ3Col))
        
        If FZBW4ch <> "NA" Then
            Set CycleB = Range(Cells(x, FZ1Col), Cells(x, FZ4Col))
        End If
        If FZBW5ch <> "NA" Then
            Set CycleB = Range(Cells(x, FZ1Col), Cells(x, FZ5Col))
        End If
        
        CycleBAvg = Application.WorksheetFunction.Average(CycleB)
        If CycleBAvg >= StabAvg - 0.5 And CycleAAvg <= StabAvg + 0.5 Then
            EndT2 = x 'this will be the end of the t2 period
            Flag = 1
        Else
            x = x + 1
        End If
    Loop


End If

'Find corresponding events on Events sheet
T2TimeOn = Cells(StartT2, 1).Value
T2TimeOff = Cells(EndT2, 1).Value
    
Sheets("Events").Activate
Cells(1, 1).Select
Selection.End(xlDown).Select
LastEvent = ActiveCell.Row

x = 1

Do
    x = x + 1
    DeltaOne = Abs(DateDiff("s", T2TimeOn, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2TimeOn, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2TimeOn = Cells(x, 1).Value
T2WHOn = Cells(x, 2).Value



x = 1
Do
    x = x + 1
    DeltaOne = Abs(DateDiff("s", T2TimeOff, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2TimeOff, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2TimeOff = Cells(x, 1).Value
T2WHOff = Cells(x, 2).Value


    
    EP2 = (T2WHOff - T2WHOn) / 1000
    T2 = Abs(DateDiff("s", T2TimeOn, T2TimeOff) / 60)
    
    If StepTwo = True Then
        MsgBox ("Unit stable before and after defrost.  Ready for Report")
        Sheets("Data Sheet").Activate
        End
    End If
    
CycleData


    'Find Ambient Averages throughout test
     Set rng = Range(Cells(IndexCycleOn(stabcycle + 1), LSOCcol), Cells(EndT2, LSOCcol))
     LSOCTotAvg = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
     Set rng = Range(Cells(IndexCycleOn(stabcycle + 1), RSOCcol), Cells(EndT2, RSOCcol))
     RSOCTotAvg = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)

End Sub




Sub CycleData()

Dim CycEnd As Integer
Dim CycStart As Integer
Dim FF1 As Range
Dim FF2 As Range
Dim FF3 As Range
Dim FF4 As Range
Dim FF5 As Range
Dim FZ1 As Range
Dim FZ2 As Range
Dim FZ3 As Range
Dim FZ4 As Range
Dim FZ5 As Range
Dim rng As Range
Dim T2LSOCHIavg(1 To 100) As Double
Dim T2LSOCLOavg(1 To 100) As Double
Dim T2RSOCHIavg(1 To 100) As Double
Dim T2RSOCLOavg(1 To 100) As Double
Dim T2Time As Date
Dim T3LSOCHIavg(1 To 100) As Double
Dim T3LSOCLOavg(1 To 100) As Double
Dim T3RSOCHIavg(1 To 100) As Double
Dim T3RSOCLOavg(1 To 100) As Double
Dim T3Time As Date
Dim x As Integer
Dim T2DefTimeOn As Integer

T2Pre = T2Pre - 1
T2Post = T2Post - 1

'Find T2 Cycle A Timestamps
Sheets("Raw Data").Activate
T2Time = Cells(T2AStart, 1).Value
x = 1
Do
    x = x + 1
    DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2On(1) = Sheets("Events").Cells(x, 1).Value
T2OffTime(1) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 1, 1), Sheets("Events").Cells(x, 1)) / 60), 2)
T2OnTime(1) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 2, 1), Sheets("Events").Cells(x + 1, 1)) / 60), 2)
T2WH(1) = Sheets("Events").Cells(x + 2, 2) - Sheets("Events").Cells(x, 2)

runtime = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 2, 1), Sheets("Events").Cells(x + 1, 1)) / 60), 2)


Sheets("Raw Data").Activate
T2Time = Cells(T2AEnd, 1).Value
x = 1
Do
    x = x + 1
    DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2Off(1) = Sheets("Events").Cells(x, 1).Value
T2Run(1) = Abs(Application.WorksheetFunction.Round((DateDiff("s", T2Off(1), T2On(1)) / 60), 2))
T2RunPer(1) = Abs(runtime / T2Run(1))

'Find T2 Precool Timestamps
Sheets("Raw Data").Activate
T2Time = Cells(T2PCStart, 1).Value
x = 1
Do
    x = x + 1
    DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2On(2) = Sheets("Events").Cells(x - 1, 1).Value
T2OffTime(2) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x - 1, 1), Sheets("Events").Cells(x, 1)) / 60), 2)
T2OnTime(2) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x, 1), Sheets("Events").Cells(x + 1, 1)) / 60), 2)
T2WH(2) = Sheets("Events").Cells(x + 1, 2) - Sheets("Events").Cells(x - 1, 2)

Sheets("Raw Data").Activate
T2Time = Cells(T2PCEnd, 1).Value
x = 1
Do
    x = x + 1
    DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2Off(2) = Sheets("Events").Cells(x, 1).Value
T2Run(2) = Abs(Application.WorksheetFunction.Round((DateDiff("s", T2On(2), T2Off(2)) / 60), 2))
runtime = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x - 1, 1), Sheets("Events").Cells(x, 1)) / 60), 2)
T2RunPer(2) = Abs(runtime / T2Run(2))

'Find T2 Defrost Timestamps
Sheets("Raw Data").Activate
T2Time = Cells(T2DefOff, 1).Value
x = 1
Do
    x = x + 1
    DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2On(3) = Sheets("Events").Cells(x, 1).Value

'Added v.1.1c
If Sheets("Events").Cells(x, 3) = 1 Then
    T2OffTime(3) = 0
Else
    T2OffTime(3) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x, 1), Sheets("Events").Cells(x + 1, 1)) / 60), 2)
End If

t2defwatton = Sheets("events").Cells(x, 2)


Sheets("Raw Data").Activate
T2Time = Cells(T2DefOn, 1).Value
x = 1
Do
    x = x + 1
    DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2DefTimeOn = x

Sheets("Raw Data").Activate
T2Time = Cells(T2DefEnd, 1).Value
x = 1
Do
    x = x + 1
    DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2OnTime(3) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(T2DefTimeOn, 1), Sheets("Events").Cells(x, 1)) / 60), 2)
T2Off(3) = Sheets("Events").Cells(x, 1).Value
T2WH(3) = Sheets("events").Cells(x, 2) - t2defwatton
T2Run(3) = Abs(Application.WorksheetFunction.Round((DateDiff("s", T2On(3), T2Off(3)) / 60), 2))
runtime = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(T2DefTimeOn, 1), Sheets("Events").Cells(x, 1)) / 60), 2)
T2RunPer(3) = Abs(runtime / T2Run(3))

'Find T2 Recovery Cycle Timestamps
Sheets("Raw Data").Activate
T2Time = Cells(T2RecStart, 1).Value

x = 1
Do
    x = x + 1
    DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2On(4) = Sheets("Events").Cells(x - 1, 1).Value
T2OnTime(4) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 1, 1), Sheets("Events").Cells(x, 1)) / 60), 2)
T2OffTime(4) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x - 1, 1), Sheets("Events").Cells(x, 1)) / 60), 2)
tempwatton = Sheets("events").Cells(x - 1, 2)

Sheets("Raw Data").Activate
T2Time = Cells(T2RecEnd, 1).Value
x = 1
Do
    x = x + 1
    DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2Off(4) = Sheets("Events").Cells(x, 1).Value
T2WH(4) = Sheets("events").Cells(x, 2) - tempwatton
T2Run(4) = Abs(Application.WorksheetFunction.Round((DateDiff("s", T2Off(4), T2On(4)) / 60), 2))
runtime = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x - 1, 1), Sheets("Events").Cells(x, 1)) / 60), 2)
T2RunPer(4) = Abs(runtime / T2Run(4))

'Find T2 B Cycle Timestamps
Sheets("Raw Data").Activate
T2Time = Cells(T2BStart, 1).Value
x = 1
Do
    x = x + 1
    If x >= (LastEvent - 2) Then
        MsgBox ("T2 period never stabilizes after defrost")
        End
    End If
    DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2On(5) = Sheets("Events").Cells(x, 1).Value
T2OffTime(5) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x, 1), Sheets("Events").Cells(x + 1, 1)) / 60), 2)
T2OnTime(5) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 1, 1), Sheets("Events").Cells(x + 2, 1)) / 60), 2)
T2WH(5) = Sheets("Events").Cells(x + 2, 2) - Sheets("Events").Cells(x, 2)
Sheets("Raw Data").Activate
T2Time = Cells(T2BEnd, 1).Value
t2runtime(5) = Sheets("events").Cells(x + 1, 1).Value
T2Run(5) = Abs(Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 2, 1), Sheets("Events").Cells(x, 1)) / 60), 2))
runtime = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 2, 1), Sheets("Events").Cells(x + 1, 1)) / 60), 2)
T2RunPer(5) = Abs(runtime / T2Run(5))
x = 1
Do
    x = x + 1
    If x >= (LastEvent - 2) Then
        MsgBox ("T2 period never stabilizes after defrost")
        End
    End If
    DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
    DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
Loop Until DeltaOne < DeltaTwo
T2Off(5) = Sheets("Events").Cells(x, 1).Value
T2BPlus = Sheets("events").Cells(x + 1, 1).Value



'Find T2 PreCycle(s) Timestamps
If T2Pre > 1 Then
    For z = 1 To T2Pre
        TempStart(z) = T2PreCycleStart(z)
        TempEnd(z) = T2PreCycleEnd(z)
    Next z
    For z = 1 To T2Pre
        zz = T2Pre + 1 - z
        T2PreCycleStart(z) = TempStart(zz)
        T2PreCycleEnd(z) = TempEnd(zz)
    Next z
End If
    
If T2Pre = 0 Then GoTo t2preend
For z = 1 To T2Pre
    
    Sheets("Raw Data").Activate
    T2Time = Cells(T2PreCycleStart(z), 1).Value
    x = 1
    Do
        x = x + 1
        DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
        DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
    Loop Until DeltaOne < DeltaTwo
    T2On(5 + z) = Sheets("Events").Cells(x, 1).Value
    T2OffTime(5 + z) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 1, 1), Sheets("Events").Cells(x, 1)) / 60), 2)
    T2OnTime(5 + z) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 2, 1), Sheets("Events").Cells(x + 1, 1)) / 60), 2)
    T2WH(5 + z) = Sheets("Events").Cells(x + 2, 2) - Sheets("Events").Cells(x, 2)
    Sheets("Raw Data").Activate

    T2Time = Cells(T2PreCycleEnd(z), 1).Value
    T2Run(5 + z) = Abs(Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 2, 1), Sheets("Events").Cells(x, 1)) / 60), 2))
    runtime = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 2, 1), Sheets("Events").Cells(x + 1, 1)) / 60), 2)
    T2RunPer(5 + z) = Abs(runtime / T2Run(5 + z))
    x = 1
    Do
        x = x + 1
        DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
        DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
    Loop Until DeltaOne < DeltaTwo
    T2Off(5 + z) = Sheets("Events").Cells(x, 1).Value
Next z
  
    
t2preend:

'Find T2 Postcycle(s) Timestamps
If T2Post = 0 Then GoTo PostEndT2
For z = 1 To T2Post
    Sheets("Raw Data").Activate
    T2Time = Cells(T2PostStart(z), 1).Value
    x = 1
    Do
        x = x + 1
        If x >= (LastEvent - 2) Then
            MsgBox ("T2 period never stabilizes after defrost")
            End
        End If
        DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
        DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
    Loop Until DeltaOne < DeltaTwo
    T2On(5 + T2Pre + z) = Sheets("Events").Cells(x, 1).Value
    T2OffTime(5 + T2Pre + z) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 1, 1), Sheets("Events").Cells(x, 1)) / 60), 2)
    T2OnTime(5 + T2Pre + z) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 2, 1), Sheets("Events").Cells(x + 1, 1)) / 60), 2)
    T2WH(5 + T2Pre + z) = Sheets("Events").Cells(x + 2, 2) - Sheets("Events").Cells(x, 2)
    T2Run(5 + T2Pre + z) = Abs(Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 2, 1), Sheets("Events").Cells(x, 1)) / 60), 2))
    runtime = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 2, 1), Sheets("Events").Cells(x + 1, 1)) / 60), 2)
    T2RunPer(5 + T2Pre + z) = Abs(runtime / T2Run(5 + T2Pre + z))
    Sheets("Raw Data").Activate
    T2Time = Cells(T2PostEnd(z), 1).Value
    x = 1
    Do
        x = x + 1
        If x >= (LastEvent - 2) Then
            MsgBox ("T2 period never stabilizes after defrost")
            End
        End If
        DeltaOne = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x, 1)) / 60)
        DeltaTwo = Abs(DateDiff("s", T2Time, Sheets("Events").Cells(x + 1, 1)) / 60)
    Loop Until DeltaOne < DeltaTwo
    T2Off(5 + T2Pre + z) = Sheets("Events").Cells(x, 1).Value
Next z
T2On(5 + T2Pre + T2Post + 1) = Sheets("events").Cells(x + 1, 1)
T2OffTime(5 + T2Pre + T2Post + 1) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 2, 1), Sheets("Events").Cells(x + 1, 1)) / 60), 2)
T2OnTime(5 + T2Pre + T2Post + 1) = Application.WorksheetFunction.Round((DateDiff("s", Sheets("Events").Cells(x + 3, 1), Sheets("Events").Cells(x + 2, 1)) / 60), 2)
T2Off(5 + T2Pre + T2Post + 1) = Sheets("events").Cells(x + 1, 1)
PostEndT2:


'Find Average Temps
Sheets("Raw Data").Activate
cycs = 5 + T2Pre + T2Post


For x = 1 To cycs
    If x = 1 Then
        CycStart = T2AStart
        CycEnd = T2AEnd
    ElseIf x = 2 Then
        CycStart = T2PCStart
        CycEnd = T2PCEnd
    ElseIf x = 3 Then
        CycStart = T2DefStart
        CycEnd = T2DefEnd
    ElseIf x = 4 Then
        CycStart = T2RecStart
        CycEnd = T2RecEnd
    ElseIf x = 5 Then
        CycStart = T2BStart
        CycEnd = T2BEnd
    ElseIf x <= (T2Pre + 5) Then
        CycStart = T2PreCycleStart(x - 5)
        CycEnd = T2PreCycleEnd(x - 5)
    Else
        CycStart = T2PostStart(x - 5 - T2Pre)
        CycEnd = T2PostEnd(x - 5 - T2Pre)
    End If
        
    Set FF1 = Range(Cells(CycStart, FFBW1col), Cells(CycEnd, FFBW1col))
    Set FF2 = Range(Cells(CycStart, FFBW2col), Cells(CycEnd, FFBW2col))
    Set FF3 = Range(Cells(CycStart, FFBW3col), Cells(CycEnd, FFBW3col))
    Set FZ1 = Range(Cells(CycStart, FZBW1col), Cells(CycEnd, FZBW1col))
    Set FZ2 = Range(Cells(CycStart, FZBW2col), Cells(CycEnd, FZBW2col))
    Set FZ3 = Range(Cells(CycStart, FZBW3col), Cells(CycEnd, FZBW3col))
 
    T2FF1(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(FF1), 2)
    T2FF2(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(FF2), 2)
    T2FF3(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(FF3), 2)
    T2FZ1(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(FZ1), 2)
    T2FZ2(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(FZ2), 2)
    T2FZ3(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(FZ3), 2)
    If FZBW4ch <> "NA" Then
        Set FZ4 = Range(Cells(CycStart, FZBW4col), Cells(CycEnd, FZBW4col))
           T2FZ4(x) = Application.WorksheetFunction.Average(FZ4)
       Else
           T2FZ4(x) = 0
       End If
       If FZBW5ch <> "NA" Then
           Set FZ5 = Range(Cells(CycStart, FZBW5col), Cells(CycEnd, FZBW5col))
           T2FZ5(x) = Application.WorksheetFunction.Average(FZ5)
       Else
           T2FZ5(x) = 0
       End If
    If ICEBWch <> "NA" Then
        Set rng = Range(Cells(CycStart, ICEBWcol), Cells(CycEnd, ICEBWcol))
        T2IB(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Else
        T2IB(x) = 0
    End If
    If Voltsch <> "NA" Then
        Set rng = Range(Cells(CycStart, Voltscol), Cells(CycEnd, Voltscol))
        T2Volts(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
        T2VoltsMin(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Min(rng), 2)
        T2VoltsMax(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Max(rng), 2)
    Else
        T2Volts(x) = 0
        T2VoltsMin(x) = 0
        T2VoltsMax(x) = 0
    End If
    
    Set rng = Range(Cells(CycStart, RSOCcol), Cells(CycEnd, RSOCcol))
    T2RsocAvg(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(CycStart, LSOCcol), Cells(CycEnd, LSOCcol))
    T2LsocAvg(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(CycStart, RSOCHIcol), Cells(CycEnd, RSOCHIcol))
    T2RSOCHIavg(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(CycStart, LSOCHIcol), Cells(CycEnd, LSOCHIcol))
    T2LSOCHIavg(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(CycStart, RSOCLOcol), Cells(CycEnd, RSOCLOcol))
    T2RSOCLOavg(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    Set rng = Range(Cells(CycStart, LSOCLOcol), Cells(CycEnd, LSOCLOcol))
    T2LSOCLOavg(x) = Application.WorksheetFunction.Round(Application.WorksheetFunction.Average(rng), 2)
    T2LGrad(x) = Abs(Application.WorksheetFunction.Round((T2LSOCHIavg(x) - T2LSOCLOavg(x)) / (hght / 12), 3))
    T2Rgrad(x) = Abs(Application.WorksheetFunction.Round((T2RSOCHIavg(x) - T2RSOCLOavg(x)) / (hght / 12), 3))

Next x

If ChooseCycleBool = False Then
        ChooseCycleAmount
    End If


End Sub
