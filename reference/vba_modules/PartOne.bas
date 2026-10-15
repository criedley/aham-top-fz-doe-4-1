Attribute VB_Name = "PartOne"
Sub EP1Calc()
Attribute EP1Calc.VB_ProcData.VB_Invoke_Func = " \n14"
    SumWattsEP1 = Cells(EndRow, WattHourscol).Value - Cells(EndRow - 179, WattHourscol).Value
    Set WattRange = Range(Cells(EndRow - 179, Wattscol), Cells(EndRow, Wattscol))
    AvgWatts = Application.WorksheetFunction.Average(WattRange)
    t1 = 180
    EP1 = SumWattsEP1 / 1000
End Sub
Sub ECalc()
Attribute ECalc.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Formula from specs at 8.7.2.1.2
    ET = (1440 * EP1 / t1) + (EP2 - (EP1 * T2 / t1)) * (12 / CT)
    If FZOnly = True Then
        ETcorr = ET * FZOnlyKfactor
    Else:
        ETcorr = ET
    End If

End Sub
