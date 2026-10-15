Attribute VB_Name = "Checks"
' Gradientcheck()
' Ambientcheck
' Stabilitycheck()
Public getlastscan As Boolean
Public FFA1R As Range
Public FFA2R As Range
Public FFA3R As Range
Public FZA1R As Range
Public FZA2R As Range
Public FZA3R As Range
Public FZA4R As Range
Public FZA5R As Range

Public FFA1 As Double
Public FFA2 As Double
Public FFA3 As Double
Public FZA1 As Double
Public FZA2 As Double
Public FZA3 As Double
Public FZA4 As Double
Public FZA5 As Double

Public FFBW1Avg As Double
Public FFBW2Avg As Double
Public FFBW3Avg As Double
Public FZBW1Avg As Double
Public FZBW2Avg As Double
Public FZBW3Avg As Double
Public FZBW4Avg As Double
Public FZBW5Avg As Double

      
Sub Gradientcheck()
Attribute Gradientcheck.VB_ProcData.VB_Invoke_Func = " \n14"
    RSOCHIavg = Application.WorksheetFunction.Average(Range(Cells(Start, RSOCHIcol), Cells(EndRow, RSOCHIcol)))
    RSOCLOavg = Application.WorksheetFunction.Average(Range(Cells(Start, RSOCLOcol), Cells(EndRow, RSOCLOcol)))
    LSOCHIavg = Application.WorksheetFunction.Average(Range(Cells(Start, LSOCHIcol), Cells(EndRow, LSOCHIcol)))
    LSOCLOavg = Application.WorksheetFunction.Average(Range(Cells(Start, LSOCLOcol), Cells(EndRow, LSOCLOcol)))
    RSGrad = (RSOCHIavg - RSOCLOavg) / (hght / 12)
    LSGrad = (LSOCHIavg - LSOCLOavg) / (hght / 12)
End Sub

Sub AmbientCheck()
Attribute AmbientCheck.VB_ProcData.VB_Invoke_Func = " \n14"
    ' changed from RowNum to EndRow to do calc on data in test cycle
    ' changed from 2 to Start to skip first compressor cycle
    RSOCMax = Application.WorksheetFunction.Max(Range(Cells(Start, RSOCcol), Cells(EndRow, RSOCcol)))
    RSOCMin = Application.WorksheetFunction.Min(Range(Cells(Start, RSOCcol), Cells(EndRow, RSOCcol)))
    LSOCMax = Application.WorksheetFunction.Max(Range(Cells(Start, LSOCcol), Cells(EndRow, LSOCcol)))
    LSOCMin = Application.WorksheetFunction.Min(Range(Cells(Start, LSOCcol), Cells(EndRow, LSOCcol)))
End Sub
Sub StabilityCheck()
Attribute StabilityCheck.VB_ProcData.VB_Invoke_Func = " \n14"

FFBW1Avg = 0
FFBW2Avg = 0
FFBW3Avg = 0
FZBW1Avg = 0
FZBW2Avg = 0
FZBW3Avg = 0
FZBW4Avg = 0
FZBW5Avg = 0

Sheets("Raw Data").Activate
getlastscan = False

    y = 0
    Do
        y = y + 1
    Loop Until Cells(1, y) = "Event 2 Defrost"
    DefrostCol = y
    
    defcheckcounter = 1
    
    If cyclic = False Then
        
        R = Start
        Delta = 300
        Stable = False
    
       Do While Stable = False And Cells(R + 420, FFBW1col) <> ""
           
           Set FFBW11 = Range(Cells(R, FFBW1col), Cells(R + 119, FFBW1col))
           Set FFBW21 = Range(Cells(R, FFBW2col), Cells(R + 119, FFBW2col))
           Set FFBW31 = Range(Cells(R, FFBW3col), Cells(R + 119, FFBW3col))
           Set FZBW11 = Range(Cells(R, FZBW1col), Cells(R + 119, FZBW1col))
           Set FZBW21 = Range(Cells(R, FZBW2col), Cells(R + 119, FZBW2col))
           Set FZBW31 = Range(Cells(R, FZBW3col), Cells(R + 119, FZBW3col))
           
           FFBW11avg = Application.WorksheetFunction.Average(FFBW11)
           FFBW21avg = Application.WorksheetFunction.Average(FFBW21)
           FFBW31avg = Application.WorksheetFunction.Average(FFBW31)
           FZBW11avg = Application.WorksheetFunction.Average(FZBW11)
           FZBW21avg = Application.WorksheetFunction.Average(FZBW21)
           FZBW31avg = Application.WorksheetFunction.Average(FZBW31)
           
           If FZBW4ch <> "NA" Then
               Set FZBW41 = Range(Cells(R, FZBW4col), Cells(R + 119, FZBW4col))
               FZBW41avg = Application.WorksheetFunction.Average(FZBW41)
           Else
               FZBW41avg = 0
           End If
           
           If FZBW5ch <> "NA" Then
               Set FZBW51 = Range(Cells(R, FZBW5col), Cells(R + 119, FZBW5col))
               FZBW51avg = Application.WorksheetFunction.Average(FZBW51)
           Else
               FZBW51avg = 0
           End If
       
           
           Set FFBW12 = Range(Cells(R + Delta, FFBW1col), Cells(R + Delta + 119, FFBW1col))
           Set FFBW22 = Range(Cells(R + Delta, FFBW2col), Cells(R + Delta + 119, FFBW2col))
           Set FFBW32 = Range(Cells(R + Delta, FFBW3col), Cells(R + Delta + 119, FFBW3col))
           Set FZBW12 = Range(Cells(R + Delta, FZBW1col), Cells(R + Delta + 119, FZBW1col))
           Set FZBW22 = Range(Cells(R + Delta, FZBW2col), Cells(R + Delta + 119, FZBW2col))
           Set FZBW32 = Range(Cells(R + Delta, FZBW3col), Cells(R + Delta + 119, FZBW3col))
           
           
               
           FFBW12avg = Application.WorksheetFunction.Average(FFBW12)
           FFBW22avg = Application.WorksheetFunction.Average(FFBW22)
           FFBW32avg = Application.WorksheetFunction.Average(FFBW32)
           FZBW12avg = Application.WorksheetFunction.Average(FZBW12)
           FZBW22avg = Application.WorksheetFunction.Average(FZBW22)
           FZBW32avg = Application.WorksheetFunction.Average(FZBW32)
               
           If FZBW4ch <> "NA" Then
               Set FZBW42 = Range(Cells(R + Delta, FZBW4col), Cells(R + Delta + 119, FZBW4col))
               FZBW42avg = Application.WorksheetFunction.Average(FZBW42)
           Else
               FZBW42avg = 0
           End If
           If FZBW5ch <> "NA" Then
               Set FZBW52 = Range(Cells(R + Delta, FZBW5col), Cells(R + Delta + 119, FZBW5col))
               FZBW52avg = Application.WorksheetFunction.Average(FZBW52)
           Else
               FZBW52avg = 0
           End If
               
           Stable = False
           Difference = (Delta / 60) * 0.042
           If (Abs(FFBW11avg - FFBW12avg) <= Difference) And Abs(FFBW21avg - FFBW22avg) <= Difference And _
              (Abs(FFBW31avg - FFBW32avg) <= Difference) And Abs(FZBW11avg - FZBW12avg) <= Difference And _
              (Abs(FZBW21avg - FZBW22avg) <= Difference) And Abs(FZBW31avg - FZBW32avg) <= Difference And _
              (Abs(FZBW41avg - FZBW42avg) <= Difference) And Abs(FZBW51avg - FZBW52avg) <= Difference Then
               
               Stable = True
               EndFLG = True
                EndRow = R + Delta + 120
               stabcycle = Application.WorksheetFunction.Round((EndRow - IndexCycleOn(1)) / 60, 1)
               If DateDiff("h", TimeCycleOn(stabcycle), Cells(EndRow, 1)) > 0 Then
                    Do While DateDiff("h", TimeCycleOn(stabcycle), Cells(EndRow, 1)) > 0
                        stabcycle = stabcycle + 1
                    Loop
               End If
                
          

                TestDuration = DateDiff("s", Cells(2, 1), Cells(EndRow, 1).Value) / 60

           End If
           R = R + 1
       Loop
       If Stable = False Then
            ExitFLG = True
            MsgBox ("test is not stable")
           EndFLG = True
           Exit Sub
       End If
       Set FFA1R = Range(Cells(EndRow - 31, FFBW1col), Cells(EndRow, FFBW1col))
       Set FFA2R = Range(Cells(EndRow - 31, FFBW2col), Cells(EndRow, FFBW2col))
       Set FFA3R = Range(Cells(EndRow - 31, FFBW3col), Cells(EndRow, FFBW3col))
       Set FZA1R = Range(Cells(EndRow - 31, FZBW1col), Cells(EndRow, FZBW1col))
       Set FZA2R = Range(Cells(EndRow - 31, FZBW2col), Cells(EndRow, FZBW2col))
       Set FZA3R = Range(Cells(EndRow - 31, FZBW3col), Cells(EndRow, FZBW3col))
 
       FFA1 = Application.WorksheetFunction.Average(FFA1R)
       FFA2 = Application.WorksheetFunction.Average(FFA2R)
       FFA3 = Application.WorksheetFunction.Average(FFA3R)
       FZA1 = Application.WorksheetFunction.Average(FZA1R)
       FZA2 = Application.WorksheetFunction.Average(FZA2R)
       FZA3 = Application.WorksheetFunction.Average(FZA3R)
       If FZBW4ch <> "NA" Then
           Set FZA4R = Range(Cells(EndRow - 31, FZBW4col), Cells(EndRow, FZBW4col))
           FZA4 = Application.WorksheetFunction.Average(FZA4R)
       Else
           FZA4 = 0
       End If
       If FZBW5ch <> "NA" Then
           Set FZA5R = Range(Cells(EndRow - 31, FZBW5col), Cells(EndRow, FZBW5col))
           FZA5 = Application.WorksheetFunction.Average(FZA5R)
       Else
           FZA5 = 0
       End If
       
    
           FFAvg = (FFA1 + FFA2 + FFA3) / 3
           If FZBW4ch <> "NA" And FZBW5ch <> "NA" Then
               FZAvg = (FZA1 + FZA2 + FZA3 + FZA4 + FZA5) / 5
           ElseIf FZBW4ch <> "NA" Then
               FZAvg = (FZA1 + FZA2 + FZA3 + FZA4) / 4
           ElseIf FZBW5ch <> "NA" Then
               FZAvg = (FZA1 + FZA2 + FZA3 + FZA4) / 4
           Else
               FZAvg = (FZA1 + FZA2 + FZA3) / 3
           End If
       
            If StepOne = True Then
                If Stable = True Then
                    MsgBox ("Unit is stable.  Ready for defrost.")
                    Sheets("Data Sheet").Activate
                    End
                Else
                    MsgBox ("Unit has not reached stability.")
                    Sheets("Data Sheet").Activate
                    End
                End If
            End If
            
       EP1Calc
       EP2calc

       ECalc
    Else
    
        ' cyclic data, another approach
        level = 1
        cyclenum = cycleCount
        Stable = False
        J = 0
        K = 1
        L = 1
        M = 1
        
ReCheckStability:
        Do While Stable = False And (J <= cyclenum)
            
            J = J + 1
            K = J
            If K <= cyclenum Then
                Do While (IndexCycleOn(K) - IndexCycleOn(J)) < 120
                    If K >= cyclenum Then
                            If NextDefrost = 0 Then
                                EndFLG = True
                                MsgBox ("[1]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                            End If
                        
                        Exit Sub
                    End If
                    K = K + 1
                Loop
            Else
                If NextDefrost = 0 Then
                    EndFLG = True
                    MsgBox ("[2]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                End If
                
                Exit Sub
            End If

            
            Set FZBW11 = Range(Cells(IndexCycleOn(J), FZBW1col), Cells(IndexCycleOn(K) - 1, FZBW1col))
            Set FZBW21 = Range(Cells(IndexCycleOn(J), FZBW2col), Cells(IndexCycleOn(K) - 1, FZBW2col))
            Set FZBW31 = Range(Cells(IndexCycleOn(J), FZBW3col), Cells(IndexCycleOn(K) - 1, FZBW3col))
    
            
            FZBW11avg = Application.WorksheetFunction.Average(FZBW11)
            FZBW21avg = Application.WorksheetFunction.Average(FZBW21)
            FZBW31avg = Application.WorksheetFunction.Average(FZBW31)
            
            If FZBW4ch <> "NA" Then
                Set FZBW41 = Range(Cells(IndexCycleOn(J), FZBW4col), Cells(IndexCycleOn(K) - 1, FZBW4col))
                FZBW41avg = Application.WorksheetFunction.Average(FZBW41)
            Else
                FZBW41avg = 0
            End If
            If FZBW5ch <> "NA" Then
                Set FZBW51 = Range(Cells(IndexCycleOn(J), FZBW5col), Cells(IndexCycleOn(K) - 1, FZBW5col))
                FZBW51avg = Application.WorksheetFunction.Average(FZBW51)
            Else
                FZBW51avg = 0
            End If
            
            Mid1 = 1440 * (TimeCycleOn(J) + (TimeCycleOff(K - 1) - TimeCycleOn(J)) / 2)
            'Find Averages for Second Stability Period
        
            L = K
            If L <= cyclenum Then
                Do While (IndexCycleOn(L) - IndexCycleOn(K)) < 180
                    If L >= cyclenum Then
                        If NextDefrost = 0 Then
                            EndFLG = True
                            MsgBox ("[3]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                        End If
                                                
                        Exit Sub
                    End If
                    L = L + 1
                Loop
            Else
                MsgBox ("[4]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                                
                Exit Sub
            End If
            
            M = L
            If M <= cyclenum Then
                Do While (IndexCycleOn(M) - IndexCycleOn(L)) < 120
                    If M >= cyclenum Then
                        If NextDefrost = 0 Then
                            EndFLG = True
                            MsgBox ("[5]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                        End If
                                                
                        Exit Sub
                    End If
                    M = M + 1
                Loop
            Else
                If NextDefrost = 0 Then
                    EndFLG = True
                    MsgBox ("[6]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                End If
                                
                Exit Sub
            End If
            
            

            Set FZBW12 = Range(Cells(IndexCycleOn(L), FZBW1col), Cells(IndexCycleOn(M) - 1, FZBW1col))
            Set FZBW22 = Range(Cells(IndexCycleOn(L), FZBW2col), Cells(IndexCycleOn(M) - 1, FZBW2col))
            Set FZBW32 = Range(Cells(IndexCycleOn(L), FZBW3col), Cells(IndexCycleOn(M) - 1, FZBW3col))
              

            FZBW12avg = Application.WorksheetFunction.Average(FZBW12)
            FZBW22avg = Application.WorksheetFunction.Average(FZBW22)
            FZBW32avg = Application.WorksheetFunction.Average(FZBW32)
                
            If FZBW4ch <> "NA" Then
                Set FZBW42 = Range(Cells(IndexCycleOn(L), FZBW4col), Cells(IndexCycleOn(M) - 1, FZBW4col))
                FZBW42avg = Application.WorksheetFunction.Average(FZBW42)
            Else
                FZBW42avg = 0
            End If
            
            If FZBW5ch <> "NA" Then
                Set FZBW52 = Range(Cells(IndexCycleOn(L), FZBW5col), Cells(IndexCycleOn(M) - 1, FZBW5col))
                FZBW52avg = Application.WorksheetFunction.Average(FZBW52)
            Else
                FZBW52avg = 0
            End If
                
            Mid2 = 1440 * (TimeCycleOn(L) + (TimeCycleOn(M - 1) - TimeCycleOn(L)) / 2)
            
            Delta = Mid2 - Mid1
            Difference = (Delta / 60) * 0.042
            
            Msg = Abs(FFBW11avg - FFBW12avg) & vbCrLf & Abs(FFBW21avg - FFBW22avg) & vbCrLf & Abs(FFBW31avg - FFBW32avg) & vbCrLf & Abs(FZBW11avg - FZBW12avg) & vbCrLf & Abs(FZBW21avg - FZBW22avg) & vbCrLf & Abs(FZBW31avg - FZBW32avg) & vbCrLf & Abs(FZBW41avg - FZBW42avg) & vbCrLf & Abs(FZBW51avg - FZBW52avg)
            
            If (Abs(FFBW11avg - FFBW12avg) <= Difference) And Abs(FFBW21avg - FFBW22avg) <= Difference And _
                (Abs(FFBW31avg - FFBW32avg) <= Difference) And Abs(FZBW11avg - FZBW12avg) <= Difference And _
                (Abs(FZBW21avg - FZBW22avg) <= Difference) And Abs(FZBW31avg - FZBW32avg) <= Difference And _
                (Abs(FZBW41avg - FZBW42avg) <= Difference) And Abs(FZBW51avg - FZBW52avg) <= Difference Then
               
                Stable = True
                EndFLG = True
            End If
            Msg = "Cycle " & (M - 1) & vbCrLf & "Stability: " & Stable & vbCrLf & "Delta: " & Difference & vbCrLf & Msg
       

        Loop

                   
Calc:
            If Stable = False Then
                ExitFLG = True
                If NextDefrost = 0 Then
                    EndFLG = True
                    MsgBox ("test is not stable")
                End If
                Exit Sub
            End If
           
         
          
            stabcycle = M - 1

           'Check for defrost before stability achieved
            For xx = defcheckcounter To stabcycle
                For defcheck = IndexCycleOn(xx) To IndexCycleOn(xx + 1)
                    If Cells(defcheck, DefrostCol) = 1 Then
                        MsgBox ("A Defrost Event occurred before unit achieved stability.")
                        J = xx
                        K = xx + 1
                        L = xx + 1
                        M = xx + 1
                        defcheckcounter = xx + 1
                        StartData = xx + 1
                        Stable = False
                        GoTo ReCheckStability
                    End If
                Next defcheck
            Next xx
            
            TestDuration = DateDiff("s", Cells(2, 1), TimeCycleOn(stabcycle + 1)) / 60

            'Find seven hour period for EP1, T1     Changed in v3.0
            J = stabcycle
            Do While (DateDiff("s", TimeCycleOn(J), TimeCycleOn(stabcycle + 1)) / 60) < 420
                J = J - 1
            Loop

            If J = stabcycle Then 'if cycle length > 180m then go to next cycle (added for 1.15)
                J = J - 1
            End If

            EP1 = (WattHoursCount(stabcycle + 1) - WattHoursCount(J)) / 1000
         
            t1 = 24 * 60 * (TimeCycleOn(stabcycle + 1) - TimeCycleOn(J))
            T1TimeOn = TimeCycleOn(J)
            StartT1 = J
            T1TimeOff = TimeCycleOn(stabcycle + 1)
            EndT1 = stabcycle
            
            If StepOne = True Then
                If Stable = True Then
                    MsgBox ("Unit is stable.  Ready for defrost.")
                    Sheets("Data Sheet").Activate
                    End
                Else
                    MsgBox ("Unit has not reached stability.")
                    Sheets("Data Sheet").Activate
                    End
                End If
            End If
            

            'Temperatures are averaged over calculated t1 and EP1 periods (v1.18)
                For I = J To stabcycle ' + 1 (removed +1 for 1.18a)
                    FFBW1Avg = FFBW1Avg + FFBW1(I)
                    FFBW2Avg = FFBW2Avg + FFBW2(I)
                    FFBW3Avg = FFBW3Avg + FFBW3(I)
                    FZBW1Avg = FZBW1Avg + FZBW1(I)
                    FZBW2Avg = FZBW2Avg + FZBW2(I)
                    FZBW3Avg = FZBW3Avg + FZBW3(I)
                    FZBW4Avg = FZBW4Avg + FZBW4(I)
                    FZBW5Avg = FZBW5Avg + FZBW5(I)
                Next I
                FFBW1Avg = FFBW1Avg / (I - J)
                FFBW2Avg = FFBW2Avg / (I - J)
                FFBW3Avg = FFBW3Avg / (I - J)
                FFAvg = (FFBW1Avg + FFBW2Avg + FFBW3Avg) / 3
                FZBW1Avg = FZBW1Avg / (I - J)
                FZBW2Avg = FZBW2Avg / (I - J)
                FZBW3Avg = FZBW3Avg / (I - J)
                If FZBW4ch <> "NA" Then
                    FZBW4Avg = FZBW4Avg / (I - J)
                End If
                If FZBW5ch <> "NA" Then
                    FZBW5Avg = FZBW5Avg / (I - J)
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
            
                EP2calc
                ET = (1440 * EP1 / t1) + ((EP2 - (EP1 * T2) / t1) * (12 / CT))
                ETcorr = ET * FZOnlyKfactor
                'FZOnlyKfactor = 1   Removed in v1.9b
            
            End If


        PrintReport
    
           
End Sub
Sub StabilityCheckFZOnly()

Dim getlastscan As Boolean
Dim FZA1R As Range
Dim FZA2R As Range
Dim FZA3R As Range
Dim FZA4R As Range
Dim FZA5R As Range

Dim FZA1 As Double
Dim FZA2 As Double
Dim FZA3 As Double
Dim FZA4 As Double
Dim FZA5 As Double

Dim FFBW1Avg As Double
Dim FFBW2Avg As Double
Dim FFBW3Avg As Double
Dim FZBW1Avg As Double
Dim FZBW2Avg As Double
Dim FZBW3Avg As Double
Dim FZBW4Avg As Double
Dim FZBW5Avg As Double

getlastscan = False

    If cyclic = False Then
        
        R = Start
        Delta = 300
        Stable = False
    
       Do While Stable = False And Cells(R + 420, FZBW1col) <> ""
           Set FZBW11 = Range(Cells(R, FZBW1col), Cells(R + 119, FZBW1col))
           Set FZBW21 = Range(Cells(R, FZBW2col), Cells(R + 119, FZBW2col))
           Set FZBW31 = Range(Cells(R, FZBW3col), Cells(R + 119, FZBW3col))
           
           FZBW11avg = Application.WorksheetFunction.Average(FZBW11)
           FZBW21avg = Application.WorksheetFunction.Average(FZBW21)
           FZBW31avg = Application.WorksheetFunction.Average(FZBW31)
           
           If FZBW4ch <> "NA" Then
               Set FZBW41 = Range(Cells(R, FZBW4col), Cells(R + 119, FZBW4col))
               FZBW41avg = Application.WorksheetFunction.Average(FZBW41)
           Else
               FZBW41avg = 0
           End If
           
           If FZBW5ch <> "NA" Then
               Set FZBW51 = Range(Cells(R, FZBW5col), Cells(R + 119, FZBW5col))
               FZBW51avg = Application.WorksheetFunction.Average(FZBW51)
           Else
               FZBW51avg = 0
           End If
           
           Set FZBW12 = Range(Cells(R + Delta, FZBW1col), Cells(R + Delta + 119, FZBW1col))
           Set FZBW22 = Range(Cells(R + Delta, FZBW2col), Cells(R + Delta + 119, FZBW2col))
           Set FZBW32 = Range(Cells(R + Delta, FZBW3col), Cells(R + Delta + 119, FZBW3col))
           
           FZBW12avg = Application.WorksheetFunction.Average(FZBW12)
           FZBW22avg = Application.WorksheetFunction.Average(FZBW22)
           FZBW32avg = Application.WorksheetFunction.Average(FZBW32)
               
           If FZBW4ch <> "NA" Then
               Set FZBW42 = Range(Cells(R + Delta, FZBW4col), Cells(R + Delta + 119, FZBW4col))
               FZBW42avg = Application.WorksheetFunction.Average(FZBW42)
           Else
               FZBW42avg = 0
           End If
           If FZBW5ch <> "NA" Then
               Set FZBW52 = Range(Cells(R + Delta, FZBW5col), Cells(R + Delta + 119, FZBW5col))
               FZBW52avg = Application.WorksheetFunction.Average(FZBW52)
           Else
               FZBW52avg = 0
           End If
               
           Stable = False
           Difference = (Delta / 60) * 0.042
           If (Abs(FZBW11avg - FZBW12avg) <= Difference) And Abs(FZBW21avg - FZBW22avg) <= Difference And _
              (Abs(FZBW31avg - FZBW32avg) <= Difference) And Abs(FZBW41avg - FZBW42avg) <= Difference And _
              (Abs(FZBW51avg - FZBW52avg) <= Difference) Then
               
               Stable = True
               EndFLG = True
                EndRow = R + Delta + 120
               stabcycle = Application.WorksheetFunction.Round((EndRow - IndexCycleOn(1)) / 60, 1)
               If DateDiff("h", TimeCycleOn(stabcycle), Cells(EndRow, 1)) > 0 Then
                    Do While DateDiff("h", TimeCycleOn(stabcycle), Cells(EndRow, 1)) > 0
                        stabcycle = stabcycle + 1
                    Loop
               End If
              

                If PreDwell = True Then
                    TestDuration = DateDiff("s", TimePreDwellOn, Cells(EndRow, 1).Value) / 60
                Else:
                    TestDuration = DateDiff("s", TimeDefOn, Cells(EndRow, 1).Value) / 60
                End If
           End If
           R = R + 1
       Loop
       If Stable = False Then
            ExitFLG = True
            MsgBox ("test is not stable")
           EndFLG = True
           Exit Sub
       End If
       Set FZA1R = Range(Cells(EndRow - 31, FZBW1col), Cells(EndRow, FZBW1col))
       Set FZA2R = Range(Cells(EndRow - 31, FZBW2col), Cells(EndRow, FZBW2col))
       Set FZA3R = Range(Cells(EndRow - 31, FZBW3col), Cells(EndRow, FZBW3col))
       
       FZA1 = Application.WorksheetFunction.Average(FZA1R)
       FZA2 = Application.WorksheetFunction.Average(FZA2R)
       FZA3 = Application.WorksheetFunction.Average(FZA3R)
       If FZBW4ch <> "NA" Then
           Set FZA4R = Range(Cells(EndRow - 31, FZBW4col), Cells(EndRow, FZBW4col))
           FZA4 = Application.WorksheetFunction.Average(FZA4R)
       Else
           FZA4 = 0
       End If
       If FZBW5ch <> "NA" Then
           Set FZA5R = Range(Cells(EndRow - 31, FZBW5col), Cells(EndRow, FZBW5col))
           FZA5 = Application.WorksheetFunction.Average(FZA5R)
       Else
           FZA5 = 0
       End If
       
    
           If FZBW4ch <> "NA" And FZBW5ch <> "NA" Then
               FZAvg = (FZA1 + FZA2 + FZA3 + FZA4 + FZA5) / 5
           ElseIf FZBW4ch <> "NA" Then
               FZAvg = (FZA1 + FZA2 + FZA3 + FZA4) / 4
           ElseIf FZBW5ch <> "NA" Then
               FZAvg = (FZA1 + FZA2 + FZA3 + FZA4) / 4
           Else
               FZAvg = (FZA1 + FZA2 + FZA3) / 3
           End If
       
       EP1Calc
       EP2calc

       ECalc
    Else
    
        ' cyclic data, another approach
        level = 1
        cyclenum = cycleCount
        Stable = False
        J = 0
        K = 1
        L = 1
        M = 1
        Do While Stable = False And (J <= cyclenum)
            J = J + 1
            K = J
            If K <= cyclenum Then
                Do While (IndexCycleOn(K) - IndexCycleOn(J)) < 120
                    If K >= cyclenum Then
                            If NextDefrost = 0 Then
                                EndFLG = True
                                MsgBox ("[1]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                            End If
                        'MsgBox (DateDiff("n", TimeCycleOn(J), TimeCycleOn(K)) & vbCrLf & DateDiff("n", TimeCycleOn(K), TimeCycleOn(L)) & vbCrLf & DateDiff("n", TimeCycleOn(L), TimeCycleOn(M)))
                        Exit Sub
                    End If
                    K = K + 1
                Loop
            Else
                If NextDefrost = 0 Then
                    EndFLG = True
                    MsgBox ("[2]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                End If
                 'MsgBox (DateDiff("n", TimeCycleOn(J), TimeCycleOn(K)) & vbCrLf & DateDiff("n", TimeCycleOn(K), TimeCycleOn(L)) & vbCrLf & DateDiff("n", TimeCycleOn(L), TimeCycleOn(M)))
                Exit Sub
            End If
    
            Set FZBW11 = Range(Cells(IndexCycleOn(J), FZBW1col), Cells(IndexCycleOn(K) - 1, FZBW1col))
            Set FZBW21 = Range(Cells(IndexCycleOn(J), FZBW2col), Cells(IndexCycleOn(K) - 1, FZBW2col))
            Set FZBW31 = Range(Cells(IndexCycleOn(J), FZBW3col), Cells(IndexCycleOn(K) - 1, FZBW3col))
            
            FZBW11avg = Application.WorksheetFunction.Average(FZBW11)
            FZBW21avg = Application.WorksheetFunction.Average(FZBW21)
            FZBW31avg = Application.WorksheetFunction.Average(FZBW31)
            
            If FZBW4ch <> "NA" Then
                Set FZBW41 = Range(Cells(IndexCycleOn(J), FZBW4col), Cells(IndexCycleOn(K) - 1, FZBW4col))
                FZBW41avg = Application.WorksheetFunction.Average(FZBW41)
            Else
                FZBW41avg = 0
            End If
            If FZBW5ch <> "NA" Then
                Set FZBW51 = Range(Cells(IndexCycleOn(J), FZBW5col), Cells(IndexCycleOn(K) - 1, FZBW5col))
                FZBW51avg = Application.WorksheetFunction.Average(FZBW51)
            Else
                FZBW51avg = 0
            End If
            
            Mid1 = 1440 * (TimeCycleOn(J) + (TimeCycleOff(K - 1) - TimeCycleOn(J)) / 2)
            'Find Averages for Second Stability Period
        
            L = K
            If L <= cyclenum Then
                Do While (IndexCycleOn(L) - IndexCycleOn(K)) < 180
                    If L >= cyclenum Then
                        If NextDefrost = 0 Then
                            EndFLG = True
                            MsgBox ("[3]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                        End If
                        'MsgBox (DateDiff("n", TimeCycleOn(J), TimeCycleOn(K)) & vbCrLf & DateDiff("n", TimeCycleOn(K), TimeCycleOn(L)) & vbCrLf & DateDiff("n", TimeCycleOn(L), TimeCycleOn(M)))
                        Exit Sub
                    End If
                    L = L + 1
                Loop
            Else
                MsgBox ("[4]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                'MsgBox (DateDiff("n", TimeCycleOn(J), TimeCycleOn(K)) & vbCrLf & DateDiff("n", TimeCycleOn(K), TimeCycleOn(L)) & vbCrLf & DateDiff("n", TimeCycleOn(L), TimeCycleOn(M)))
                Exit Sub
            End If
            
            M = L
            If M <= cyclenum Then
                Do While (IndexCycleOn(M) - IndexCycleOn(L)) < 120
                    If M >= cyclenum Then
                        If NextDefrost = 0 Then
                            EndFLG = True
                            MsgBox ("[5]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                        End If
                        'MsgBox (DateDiff("n", TimeCycleOn(J), TimeCycleOn(K)) & vbCrLf & DateDiff("n", TimeCycleOn(K), TimeCycleOn(L)) & vbCrLf & DateDiff("n", TimeCycleOn(L), TimeCycleOn(M)))
                        Exit Sub
                    End If
                    M = M + 1
                Loop
            Else
                If NextDefrost = 0 Then
                    EndFLG = True
                    MsgBox ("[6]:test not stable, aborting J:" & J & " K:" & K & " L:" & L & " M:" & M)
                End If
                'MsgBox (DateDiff("n", TimeCycleOn(J), TimeCycleOn(K)) & vbCrLf & DateDiff("n", TimeCycleOn(K), TimeCycleOn(L)) & vbCrLf & DateDiff("n", TimeCycleOn(L), TimeCycleOn(M)))
                Exit Sub
            End If
                       
            Set FZBW12 = Range(Cells(IndexCycleOn(L), FZBW1col), Cells(IndexCycleOn(M) - 1, FZBW1col))
            Set FZBW22 = Range(Cells(IndexCycleOn(L), FZBW2col), Cells(IndexCycleOn(M) - 1, FZBW2col))
            Set FZBW32 = Range(Cells(IndexCycleOn(L), FZBW3col), Cells(IndexCycleOn(M) - 1, FZBW3col))
            
            FZBW12avg = Application.WorksheetFunction.Average(FZBW12)
            FZBW22avg = Application.WorksheetFunction.Average(FZBW22)
            FZBW32avg = Application.WorksheetFunction.Average(FZBW32)
                
            If FZBW4ch <> "NA" Then
                Set FZBW42 = Range(Cells(IndexCycleOn(L), FZBW4col), Cells(IndexCycleOn(M) - 1, FZBW4col))
                FZBW42avg = Application.WorksheetFunction.Average(FZBW42)
            Else
                FZBW42avg = 0
            End If
            
            If FZBW5ch <> "NA" Then
                Set FZBW52 = Range(Cells(IndexCycleOn(L), FZBW5col), Cells(IndexCycleOn(M) - 1, FZBW5col))
                FZBW52avg = Application.WorksheetFunction.Average(FZBW52)
            Else
                FZBW52avg = 0
            End If
                
            Mid2 = 1440 * (TimeCycleOn(L) + (TimeCycleOn(M - 1) - TimeCycleOn(L)) / 2)
            
            Delta = Mid2 - Mid1
            Difference = (Delta / 60) * 0.042
            
            Msg = Abs(FZBW11avg - FZBW12avg) & vbCrLf & Abs(FZBW21avg - FZBW22avg) & vbCrLf & Abs(FZBW31avg - FZBW32avg) & vbCrLf & Abs(FZBW41avg - FZBW42avg) & vbCrLf & Abs(FZBW51avg - FZBW52avg)
            
             If (Abs(FZBW11avg - FZBW12avg) <= Difference) And Abs(FZBW21avg - FZBW22avg) <= Difference And _
                (Abs(FZBW31avg - FZBW32avg) <= Difference) And Abs(FZBW41avg - FZBW42avg) <= Difference And _
                (Abs(FZBW51avg - FZBW52avg) <= Difference) Then
               
                Stable = True
                EndFLG = True
            End If
            Msg = "Cycle " & (M - 1) & vbCrLf & "Stability: " & Stable & vbCrLf & "Delta: " & Difference & vbCrLf & Msg
          
        Loop
            
Calc:
            If Stable = False Then
                ExitFLG = True
                If NextDefrost = 0 Then
                    EndFLG = True
                    MsgBox ("test is not stable")
                End If
                Exit Sub
            End If
            stabcycle = M - 1
            'TestDuration = DateDiff("s", TimeDefOn, TimeCycleOn(stabcycle + 1)) / 60
            If PreDwell = True Then
             TestDuration = DateDiff("s", TimePreDwellOn, TimeCycleOn(stabcycle + 1)) / 60
            Else:
             TestDuration = DateDiff("s", TimeDefOn, TimeCycleOn(stabcycle + 1)) / 60
            End If
            'Find seven hour period for EP1, T1     Changed in v3.0
            J = stabcycle
            Do While (DateDiff("s", TimeCycleOn(J), TimeCycleOn(stabcycle + 1)) / 60) < 420
                J = J - 1
            Loop
            
            EP1 = (WattHoursCount(stabcycle + 1) - WattHoursCount(J)) / 1000
            t1 = 24 * 60 * (TimeCycleOn(stabcycle + 1) - TimeCycleOn(J))
            StartT1 = J
            EndT1 = stabcycle + 1
            EP2calc
       
                      
            'Temperatures are averaged over calculated t1 and EP1 periods (v1.18)
            For I = J To stabcycle ' + 1 (removed +1 for 1.18a)
                FFBW1Avg = FFBW1Avg + FFBW1(I)
                FFBW2Avg = FFBW2Avg + FFBW2(I)
                FFBW3Avg = FFBW3Avg + FFBW3(I)
                FZBW1Avg = FZBW1Avg + FZBW1(I)
                FZBW2Avg = FZBW2Avg + FZBW2(I)
                FZBW3Avg = FZBW3Avg + FZBW3(I)
                FZBW4Avg = FZBW4Avg + FZBW4(I)
                FZBW5Avg = FZBW5Avg + FZBW5(I)
            Next I
            FZBW1Avg = FZBW1Avg / (I - J)
            FZBW2Avg = FZBW2Avg / (I - J)
            FZBW3Avg = FZBW3Avg / (I - J)
            If FZBW4ch <> "NA" Then
                FZBW4Avg = FZBW4Avg / (I - J)
            End If
            If FZBW5ch <> "NA" Then
                FZBW5Avg = FZBW5Avg / (I - J)
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


            ET = (1440 * EP1 / t1) + ((EP2 - (EP1 * T2) / t1) * (12 / CT))
            If FZOnly = True Then
                ETcorr = ET * FZOnlyKfactor
            Else:
                ETcorr = ET * FZOnlyKfactor
            End If
        End If

        PrintReport
    
           
End Sub




