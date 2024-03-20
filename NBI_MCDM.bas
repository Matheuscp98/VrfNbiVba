Attribute VB_Name = "NBI_Completed"
' Enable full screen mode and hide interface elements
Sub EnableFullScreen()

    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    
End Sub

' Disable full screen mode and restore interface elements
Sub DisableFullScreen()

    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    
End Sub

' Save the workbook while temporarily turning off screen updating
Sub SaveWorkbook()

    Application.ScreenUpdating = False

    ThisWorkbook.Save

    Application.ScreenUpdating = True

End Sub

' Clear specific ranges and reset Solver settings
Sub ClearCells3()

    ' Clear Solver result ranges
    Range("Z3:AM68").ClearContents
    Range("Z74:AM139").ClearContents
    Range("Z145:AM210").ClearContents
    
    ' Clear variable ranges
    Range("M4:M6").ClearContents
    
    ' Clear constraint ranges
    Range("AX3:BA68").ClearContents
    Range("AX74:BA139").ClearContents
    Range("AX145:BA210").ClearContents
    
    ' Reset 'n' to its initial value
    Range("T14").Value = 1
    
    ' Reset Solver settings
    SolverReset

    ' Return to cell A1
    Range("A1").Select
    
End Sub

' Clear specific ranges and reset Solver settings
Sub ClearCells8()

    ' Clear Solver result ranges
    Range("AG3:AQ794").ClearContents
    Range("AG799:AQ1590").ClearContents
    Range("AG1595:AQ2386").ClearContents
    
    ' Clear variable ranges
    Range("J4:J6").ClearContents
    
    ' Clear constraint ranges
    Range("BA3:BI794").ClearContents
    Range("BA799:BI1590").ClearContents
    Range("BA1595:BI2386").ClearContents
    
    ' Reset 'n' to its initial value
    Range("V23").Value = 1
    
    ' Reset Solver settings
    SolverReset

    ' Return to cell A1
    Range("A1").Select
    
End Sub

' Clear specific ranges and reset Solver settings
Sub ClearCellsPost()

    ' Clear Solver result ranges
    Range("M3:W23").ClearContents
    
    ' Clear variable ranges
    Range("G3:G5").ClearContents
    
    ' Reset 'n' and 'BetaMD' to their initial values
    Range("E27").Value = 1
    Range("C33").Value = 1
    
    ' Reset Solver settings
    SolverReset

    ' Return to cell A1
    Range("A1").Select
    
    
End Sub

' Clear data points from a specified range
Sub ClearDataPoints3()

    ' Clear data points range
    Range("DA7:EB10000").ClearContents
    Range("DA3").Select

End Sub

' Clear data points from a specified range
Sub ClearDataPoints8()

    ' Clear data points range
    Range("DI7:EQ10000").ClearContents
    Range("DI3").Select

End Sub

' Solving VRF-NBI while retaining previous x data points
Sub NBIASolve()

    ' Reset all Solver settings
    SolverReset
    
    ' Loop from 1 to 66 for n
    For i = 1 To 66
        ' Set the value of n
        Range("$T$14").Value = i
        
        
        ' Execute the Solver
        SolverOk SetCell:="$C$17", MaxMinVal:=1, ValueOf:=0, ByChange:="$M$4:$M$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$16"
        SolverAdd CellRef:="$E$15:$E$17", Relation:=2, FormulaText:="0"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        
        ' Copy and paste the results
        Range("$C$32:$P$32").Copy
        Range("Z" & (i + 73) & ":AM" & (i + 73)).PasteSpecial Paste:=xlPasteValues
        
        Range("C15").Copy
        Range("AX" & (i + 73)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("E15:E17").Copy
        Range("AY" & (i + 73) & ":BA" & (i + 73)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Application.CutCopyMode = False
        
        ' Reset all Solver settings
        SolverReset
        
    Next i
    
    Range("T14").Value = 1
    Range("T14").Select
    Range("M4:M6").ClearContents
    
    ' Adjust borders
    Range("V74:AM139").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("V74:AM139").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        Range("AX74:BA139").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AX74:BA139").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    MsgBox "VRF-NBI completed with previous points retained for each iteration!"
    Range("A1").Select

End Sub

' Solving VRF-NBI with x points being an average of the optimal points
Sub NBIOSolve()

    ' Reset all Solver settings
    SolverReset

    ' Loop from 1 to 66 for n
    For i = 1 To 66
        ' Set the value of n
        Range("$T$14").Value = i
        
        Range("M4").Value = Range("Z69").Value
        Range("M5").Value = Range("AA69").Value
        Range("M6").Value = Range("AB69").Value
        
        ' Execute the Solver
        SolverOk SetCell:="$C$17", MaxMinVal:=1, ValueOf:=0, ByChange:="$M$4:$M$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$16"
        SolverAdd CellRef:="$E$15:$E$17", Relation:=2, FormulaText:="0"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        
        ' Copy and paste the results
        Range("$C$32:$P$32").Copy
        Range("Z" & (i + 144) & ":AM" & (i + 144)).PasteSpecial Paste:=xlPasteValues
        Range("C15").Copy
        Range("AX" & (i + 144)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("E15:E17").Copy
        Range("AY" & (i + 144) & ":BA" & (i + 144)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Application.CutCopyMode = False
        
        ' Reset all Solver settings
        SolverReset
        
    Next i
    
    
    Range("T14").Value = 1
    Range("T14").Select
    Range("M4:M6").ClearContents
    
    ' Adjust borders
    Range("V145:AM210").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("V145:AM210").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        Range("AX145:BA210").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AX145:BA210").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    MsgBox "VRF-NBI completed with the optimal points retained for each iteration!"
    Range("A1").Select

End Sub

' Solving VRF-NBI with x points always reset to zero
Sub NBIZSolve()

    ' Reset all Solver settings
    SolverReset

    ' Loop from 1 to 66 for n
    For i = 1 To 66
        ' Set the value of n
        Range("$T$14").Value = i
        
        Range("M4").Value = 0
        Range("M5").Value = 0
        Range("M6").Value = 0
        
        ' Execute the Solver
        SolverOk SetCell:="$C$17", MaxMinVal:=1, ValueOf:=0, ByChange:="$M$4:$M$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$16"
        SolverAdd CellRef:="$E$15:$E$17", Relation:=2, FormulaText:="0"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        
        ' Copy and paste the results
        Range("$C$32:$P$32").Copy
        Range("Z" & (i + 2) & ":AM" & (i + 2)).PasteSpecial Paste:=xlPasteValues
        Range("C15").Copy
        Range("AX" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("E15:E17").Copy
        Range("AY" & (i + 2) & ":BA" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Application.CutCopyMode = False
        
        ' Reset all Solver settings
        SolverReset
        
    Next i
    
    
    Range("T14").Value = 1
    Range("T14").Select
    Range("M4:M6").ClearContents
    
    ' Adjust borders
    Range("V3:AM68").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("V3:AM68").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        Range("AX3:BA68").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AX3:BA68").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    MsgBox "VRF-NBI completed with the previous points reset to zero for each iteration!"
    Range("A1").Select

End Sub

' Solving NBI-8Y with  retaining previous x data points
Sub NBIA8Solve()

    ' Reset all Solver settings
    SolverReset
    
    ' Loop from 1 to 792 for n
    For i = 1 To 792
        ' Set the value of n
        Range("$V$23").Value = i
        
        ' Execute the Solver
        SolverOk SetCell:="$C$17", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$20"
        SolverAdd CellRef:="$E$15:$E$17", Relation:=2, FormulaText:="0"
        SolverAdd CellRef:="$G$15:$G$17", Relation:=2, FormulaText:="0"
        SolverAdd CellRef:="$I$15:$I$16", Relation:=2, FormulaText:="0"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        
        ' Copy and paste the results
        Range("$C$33:$M$33").Copy
        Range("AG" & (i + 798) & ":AQ" & (i + 798)).PasteSpecial Paste:=xlPasteValues
        
        Range("C15").Copy
        Range("BA" & (i + 798)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("E15:E17").Copy
        Range("BB" & (i + 798) & ":BD" & (i + 798)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("G15:G17").Copy
        Range("BE" & (i + 798) & ":BG" & (i + 798)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("I15:I16").Copy
        Range("BH" & (i + 798) & ":BI" & (i + 798)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Application.CutCopyMode = False
        
        ' Reset all Solver settings
        SolverReset
        
    Next i
    
    
    Range("V23").Value = 1
    Range("V23").Select
    Range("J4:J6").ClearContents
    
    ' Adjust borders
    Range("AG3:AQ794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AG3:AQ794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        Range("BA3:BI794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("BA3:BI794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    MsgBox "NBI-8Y completed with previous points retained for each iteration!"
    Range("A1").Select

End Sub

'Solving VRF-NBINBI with x points being an average of the optimal points.
Sub NBIO8Solve()

    ' Adjust borders
    SolverReset

    ' Loop from 1 to 792 for n
    For i = 1 To 792
        ' Set the value of n
        Range("$V$23").Value = i
        
        Range("J4").Value = Range("AG797").Value
        Range("J5").Value = Range("AH797").Value
        Range("J6").Value = Range("AI797").Value
        
        ' Execute the Solver
        SolverOk SetCell:="$C$17", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$20"
        SolverAdd CellRef:="$E$15:$E$17", Relation:=2, FormulaText:="0"
        SolverAdd CellRef:="$G$15:$G$17", Relation:=2, FormulaText:="0"
        SolverAdd CellRef:="$I$15:$I$16", Relation:=2, FormulaText:="0"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        
        ' Copy and paste the results
        Range("$C$33:$M$33").Copy
        Range("AG" & (i + 1594) & ":AQ" & (i + 1594)).PasteSpecial Paste:=xlPasteValues
        
        Range("C15").Copy
        Range("BA" & (i + 1594)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("E15:E17").Copy
        Range("BB" & (i + 1594) & ":BD" & (i + 1594)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("G15:G17").Copy
        Range("BE" & (i + 1594) & ":BG" & (i + 1594)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("I15:I16").Copy
        Range("BH" & (i + 1594) & ":BI" & (i + 1594)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Application.CutCopyMode = False
        
        ' Reset all Solver settings
        SolverReset
        
    Next i
    
    
    Range("V23").Value = 1
    Range("V23").Select
    Range("J4:J6").ClearContents
    
    ' Adjust borders
    Range("AG3:AQ794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AG3:AQ794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        Range("BA3:BI794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("BA3:BI794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    MsgBox "NBI-8Y completed with the optimal points retained for each iteration!"
    Range("A1").Select

End Sub

' Solving NBI-8Y with x points always reset to zero
Sub NBIZ8Solve()

    ' Reset all Solver settings
    SolverReset

    ' Loop from 1 to 792 for n
    For i = 1 To 792
        ' Set the value of n
        Range("$V$23").Value = i
        
        Range("J4").Value = 0
        Range("J5").Value = 0
        Range("J6").Value = 0
        
        ' Execute the Solver
        SolverOk SetCell:="$C$17", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$20"
        SolverAdd CellRef:="$E$15:$E$17", Relation:=2, FormulaText:="0"
        SolverAdd CellRef:="$G$15:$G$17", Relation:=2, FormulaText:="0"
        SolverAdd CellRef:="$I$15:$I$16", Relation:=2, FormulaText:="0"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        
        ' Copy and paste the results
        Range("$C$33:$M$33").Copy
        Range("AG" & (i + 2) & ":AQ" & (i + 2)).PasteSpecial Paste:=xlPasteValues
        
        Range("C15").Copy
        Range("BA" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("E15:E17").Copy
        Range("BB" & (i + 2) & ":BD" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("G15:G17").Copy
        Range("BE" & (i + 2) & ":BG" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("I15:I16").Copy
        Range("BH" & (i + 2) & ":BI" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Application.CutCopyMode = False
        
        ' Reset all Solver settings
        SolverReset
        
    Next i
    
    Range("V23").Value = 1
    Range("V23").Select
    Range("J4:J6").ClearContents
    
    ' Adjust borders
    Range("AG3:AQ794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AG3:AQ794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        Range("BA3:BI794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("BA3:BI794").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    MsgBox "NBI-8Y completed with the previous points reset to zero for each iteration!"
    Range("A1").Select

End Sub

' NBI post-optimization while retaining previous points
Sub NBIAPost()

    ' Reset all Solver settings
    SolverReset
    
    ' Loop from 1 to 21 for n
    For i = 1 To 21
    ' Set the value of n
    Range("$C$33").Value = i
    
        ' Execute the Solver
        SolverOk SetCell:="$C$27", MaxMinVal:=2, ValueOf:=0, ByChange:="$G$3:$G$5", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$20", Relation:=1, FormulaText:="1"
        SolverAdd CellRef:="$E$20", Relation:=2, FormulaText:="1"
        SolverAdd CellRef:="$G$3:$G$5", Relation:=1, FormulaText:="1"
        SolverAdd CellRef:="$G$3:$G$5", Relation:=3, FormulaText:="0"
        SolverAdd CellRef:="$G$27", Relation:=3, FormulaText:="$G$28"
        SolverSolve True
        
        ' Copy and paste the results
        Range("$B$31:$J$31").Copy
        Range("L" & (i + 2) & ":T" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("C20").Copy
        Range("V" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("E20").Copy
        Range("W" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Application.CutCopyMode = False


        ' Reset all Solver settings
        SolverReset
        
        ' Set Beta
        Range("E27").Value = Range("E27").Value - 0.05
        
    Next i
    
    Range("E27").Value = 1
    Range("C33").Value = 1
    Range("G3:G5").ClearContents
    
    MsgBox "NBI pos-optimization completed with previous points retained for each iteration!"
    Range("A1").Select


End Sub

' NBI post-optimization by resetting the points.
Sub NBIZPost()

    ' Reset all Solver settings
    SolverReset
    
    ' Loop from 1 to 21 for n
    For i = 1 To 21
    ' Set the value of n
    Range("$C$33").Value = i
    
    Range("G3").Value = 0
    Range("G4").Value = 0
    Range("G5").Value = 0
    
        ' Execute the Solver
        SolverOk SetCell:="$C$27", MaxMinVal:=2, ValueOf:=0, ByChange:="$G$3:$G$5", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$20", Relation:=1, FormulaText:="1"
        SolverAdd CellRef:="$E$20", Relation:=2, FormulaText:="1"
        SolverAdd CellRef:="$G$3:$G$5", Relation:=1, FormulaText:="1"
        SolverAdd CellRef:="$G$3:$G$5", Relation:=3, FormulaText:="0"
        SolverAdd CellRef:="$G$27", Relation:=3, FormulaText:="$G$28"
        SolverSolve True
        
        ' Copy and paste the results
        Range("$B$31:$J$31").Copy
        Range("L" & (i + 2) & ":T" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("C20").Copy
        Range("V" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("E20").Copy
        Range("W" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Application.CutCopyMode = False


        ' Reset all Solver settings
        SolverReset
        
        ' Set Beta
        Range("E27").Value = Range("E27").Value - 0.05
        
    Next i
    
    Range("E27").Value = 1
    Range("C33").Value = 1
    Range("G3:G5").ClearContents
    
    MsgBox "NBI pos-optimization completed with the previous points reset to zero for each iteration!"
    Range("A1").Select


End Sub

' Solve NBI post-optimization RSM
Sub NBIPostRSM()

    ' Execute the Solver
    SolverOk SetCell:="$C$17", MaxMinVal:=1, ValueOf:=0, ByChange:="$M$4:$M$6", Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
    0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
    SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
    :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
    IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
    SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$16"
    SolverAdd CellRef:="$E$15:$E$17", Relation:=2, FormulaText:="0"
    SolverOptions AssumeNonNeg:=False
    SolverSolve True
        
    ' Copy and paste the results
    Range("$B$13:$L$13").Copy
    Range("$C$38:$M$38").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False

    ' Adjust borders
    SolverReset

    ' Select cell with decoded x values
    Range("U22").Select
    
End Sub

' Individual optimization for 3 factors (VRF)
Sub OptiInd3()

    ' Reset all Solver settings
    SolverReset
    
    ' Check the value to determine Max or Min
    If Range("C27").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$J$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$M$4:$M$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C27").Select
    
    ElseIf Range("C27").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$J$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$M$4:$M$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C27").Select
        
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("J13:L13").Copy
    Range("O3:O5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("J13").Select
    Range("M4:M6").ClearContents

    ' Reset all Solver settings
    SolverReset
    
    ' Check the value to determine Max or Min
    If Range("C28").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$K$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$M$4:$M$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C28").Select
    
    ElseIf Range("C28").Value = "Minimmization" Then
        ' Execute the Solver
        SolverOk SetCell:="$K$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$M$4:$M$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C28").Select
        
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("J13:L13").Copy
    Range("P3:P5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("K13").Select
    Range("M4:M6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C29").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$L$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$M$4:$M$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C29").Select
        
    
    ElseIf Range("C29").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$L$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$M$4:$M$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C29").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("J13:L13").Copy
    Range("Q3:Q5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("L13").Select
    Range("M4:M6").ClearContents
    
    Application.CutCopyMode = False
    Range("A1").Select
    
    ' Check and set the formula
    If Range("C27").Value = "Max" Then
        Range("S3").FormulaLocal = "=MAX(O3:Q3)"
        Range("T3").FormulaLocal = "=MIN(O3:Q3)"
        Range("AV3").FormulaLocal = "=MAX(AK3:AK68)"
    ElseIf Range("C27").Value = "Min" Then
        Range("S3").FormulaLocal = "=MIN(O3:Q3)"
        Range("T3").FormulaLocal = "=MAX(O3:Q3)"
        Range("AV3").FormulaLocal = "=MIN(AK3:AK68)"
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
    End If
    
    ' Check and set the formula
    If Range("C28").Value = "Max" Then
        Range("S4").FormulaLocal = "=MAX(O4:Q4)"
        Range("T4").FormulaLocal = "=MIN(O4:Q4)"
        Range("AV4").FormulaLocal = "=MAX(AL3:AL68)"
    ElseIf Range("C28").Value = "Min" Then
        Range("S4").FormulaLocal = "=MIN(O4:Q4)"
        Range("T4").FormulaLocal = "=MAX(O4:Q4)"
        Range("AV4").FormulaLocal = "=MIN(AL3:AL68)"
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
    End If
    
    ' Check and set the formula
    If Range("C29").Value = "MMaxo" Then
        Range("S5").FormulaLocal = "=MAX(O5:Q5)"
        Range("T5").FormulaLocal = "=MIN(O5:Q5)"
        Range("AV5").FormulaLocal = "=MAX(AM3:AM68)"
    ElseIf Range("C29").Value = "Min" Then
        Range("S5").FormulaLocal = "=MIN(O5:Q5)"
        Range("T5").FormulaLocal = "=MAX(O5:Q5)"
        Range("AV5").FormulaLocal = "=MIN(AM3:AM68)"
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
    End If

    ' Replace "@" with "=" in the formulas
    Range("AV3:AV5").Replace What:="@", Replacement:="", LookAt:=xlPart
    
   ' Reset all Solver settings
    SolverReset

End Sub

' Individual optimization for 8 original variables
Sub OptiInd8()

    ' Reset all Solver settings
    SolverReset
    
    ' Check the value to determine Max or Min
    If Range("C19").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$B$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C19").Select
    
    ElseIf Range("C19").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$B$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C19").Select
        
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("L3:L10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("B13").Select
    Range("J4:J6").ClearContents

    ' Reset all Solver settings
    SolverReset
    
    ' Check the value to determine Max or Min
    If Range("C20").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$C$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C20").Select
    
    ElseIf Range("C20").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$C$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C20").Select
        
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("M3:M10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("C13").Select
    Range("J4:J6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C21").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$D$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C21").Select
        
    
    ElseIf Range("C21").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$D$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C21").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("N3:N10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("D13").Select
    Range("J4:J6").ClearContents

    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C22").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$E$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C22").Select
        
    
    ElseIf Range("C22").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$E$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C22").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("O3:O10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("E13").Select
    Range("J4:J6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C23").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$F$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C23").Select
        
    
    ElseIf Range("C23").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$F$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C23").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("P3:P10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("F13").Select
    Range("J4:J6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C24").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$G$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C24").Select
        
    
    ElseIf Range("C24").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$G$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C24").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("Q3:Q10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("G13").Select
    Range("J4:J6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C25").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$H$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C25").Select
        
    
    ElseIf Range("C25").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$H$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C25").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("R3:R10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("H13").Select
    Range("J4:J6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C26").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$I$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C26").Select
        
    
    ElseIf Range("C26").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$I$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C26").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("S3:S10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("I13").Select
    Range("J4:J6").ClearContents
    

    Application.CutCopyMode = False
    Range("A1").Select
    
    ' Check and set the formula
    If Range("C19").Value = "Max" Then
        Range("U3").FormulaLocal = "=MAX(L3:S3)"
        Range("V3").FormulaLocal = "=MIN(L3:S3)"
        Range("AY3").FormulaLocal = "=MAX(AJ3:AJ794)"
    ElseIf Range("C19").Value = "Min" Then
        Range("U3").FormulaLocal = "=MIN(L3:S3)"
        Range("V3").FormulaLocal = "=MAX(L3:S3)"
        Range("AY3").FormulaLocal = "=MIN(AJ3:AJ794)"
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
    End If
    
    ' Check and set the formula
    If Range("C20").Value = "Max" Then
        Range("U4").FormulaLocal = "=MAX(L4:S4)"
        Range("V4").FormulaLocal = "=MIN(L4:S4)"
        Range("AY4").FormulaLocal = "=MAX(AK3:AK794)"
    ElseIf Range("C20").Value = "Min" Then
        Range("U4").FormulaLocal = "=MIN(L4:S4)"
        Range("V4").FormulaLocal = "=MAX(L4:S4)"
        Range("AY4").FormulaLocal = "=MIN(AK3:AK794)"
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
    End If
    
    ' Check and set the formula
    If Range("C21").Value = "Max" Then
        Range("U3").FormulaLocal = "=MAX(L5:S5)"
        Range("V5").FormulaLocal = "=MIN(L5:S5)"
        Range("AY5").FormulaLocal = "=MAX(AL3:AL794)"
    ElseIf Range("C21").Value = "Min" Then
        Range("U5").FormulaLocal = "=MIN(L5:S5)"
        Range("V5").FormulaLocal = "=MAX(L5:S5)"
        Range("AY5").FormulaLocal = "=MIN(AL3:AL794)"
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
    End If
    
    ' Check and set the formula
    If Range("C22").Value = "Max" Then
        Range("U6").FormulaLocal = "=MAX(L6:S6)"
        Range("V6").FormulaLocal = "=MIN(L6:S6)"
        Range("AY6").FormulaLocal = "=MAX(AM3:AM794)"
    ElseIf Range("C22").Value = "Min" Then
        Range("U6").FormulaLocal = "=MIN(L6:S6)"
        Range("V6").FormulaLocal = "=MAX(L6:S6)"
        Range("AY6").FormulaLocal = "=MIN(AM3:AM794)"
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
    End If
    
    ' Check and set the formula
    If Range("C23").Value = "Max" Then
        Range("U7").FormulaLocal = "=MAX(L7:S7)"
        Range("V7").FormulaLocal = "=MIN(L7:S7)"
        Range("AY7").FormulaLocal = "=MAX(AN3:AN794)"
    ElseIf Range("C23").Value = "Min" Then
        Range("U7").FormulaLocal = "=MIN(L7:S7)"
        Range("V7").FormulaLocal = "=MAX(L7:S7)"
        Range("AY7").FormulaLocal = "=MIN(AN3:AN794)"
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
    End If
    
    ' Check and set the formula
    If Range("C24").Value = "Max" Then
        Range("U8").FormulaLocal = "=MAX(L8:S8)"
        Range("V8").FormulaLocal = "=MIN(L8:S8)"
        Range("AY8").FormulaLocal = "=MAX(AO3:AO794)"
    ElseIf Range("C24").Value = "Min" Then
        Range("U8").FormulaLocal = "=MIN(L8:S8)"
        Range("V8").FormulaLocal = "=MAX(L8:S8)"
        Range("AY8").FormulaLocal = "=MIN(AO3:AO794)"
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
    End If
    
    ' Check and set the formula
    If Range("C25").Value = "Max" Then
        Range("U9").FormulaLocal = "=MAX(L9:S9)"
        Range("V9").FormulaLocal = "=MIN(L9:S9)"
        Range("AY9").FormulaLocal = "=MAX(AP3:AP794)"
    ElseIf Range("C25").Value = "Min" Then
        Range("U9").FormulaLocal = "=MIN(L9:S9)"
        Range("V9").FormulaLocal = "=MAX(L9:S9)"
        Range("AY9").FormulaLocal = "=MIN(AP3:AP794)"
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
    End If
    
    ' Check and set the formula
    If Range("C26").Value = "Max" Then
        Range("U10").FormulaLocal = "=MAX(L10:S10)"
        Range("V10").FormulaLocal = "=MIN(L10:S10)"
        Range("AY10").FormulaLocal = "=MAX(AQ3:AQ794)"
    ElseIf Range("C26").Value = "Min" Then
        Range("U10").FormulaLocal = "=MIN(L10:S10)"
        Range("V10").FormulaLocal = "=MAX(L10:S10)"
        Range("AY10").FormulaLocal = "=MIN(AQ3:AQ794)"
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
    End If
    
    ' Replace "@" with "=" in the formulas
    Range("AY3:AY10").Replace What:="@", Replacement:="", LookAt:=xlPart
    
    ' Reset all Solver settings
    SolverReset

End Sub

' Individual optimization for post-optimization VRF-NBI
Sub OptiIndPost()

    ' Reset all Solver settings
    SolverReset
    
    ' Execute the Solver
    SolverOk SetCell:="$C$18", MaxMinVal:=2, ValueOf:=0, ByChange:="$G$3:$G$5", Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
    0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
    SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
    :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
    IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
    SolverAdd CellRef:="$C$20", Relation:=1, FormulaText:="1"
    SolverAdd CellRef:="$E$20", Relation:=2, FormulaText:="1"
    SolverAdd CellRef:="$G$3:$G$5", Relation:=1, FormulaText:="1"
    SolverAdd CellRef:="$G$3:$G$5", Relation:=3, FormulaText:="0"
    SolverSolve True
    Range("C18").Select
    
    ' Copy and paste
    Range("C18:D18").Copy
    Range("I3:I4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("E18:F18").Copy
    Range("I7:I8").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("G3:G5").ClearContents

    ' Reset all Solver settings
    SolverReset
    
    ' Execute the Solver
    SolverOk SetCell:="$D$18", MaxMinVal:=1, ValueOf:=0, ByChange:="$G$3:$G$5", Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
    0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
    SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
    :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
    IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
    SolverAdd CellRef:="$C$20", Relation:=1, FormulaText:="1"
    SolverAdd CellRef:="$E$20", Relation:=2, FormulaText:="1"
    SolverAdd CellRef:="$G$3:$G$5", Relation:=1, FormulaText:="1"
    SolverAdd CellRef:="$G$3:$G$5", Relation:=3, FormulaText:="0"
    SolverSolve True
    Range("D18").Select
    
    ' Copy and paste
    Range("C18:D18").Copy
    Range("J3:J4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("E18:F18").Copy
    Range("J7:J8").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("G3:G5").ClearContents

    SolverReset
    
    Range("A1").Select
    
    
End Sub

' Calculate the Payoff matrix 3Y
Sub PayoffMatrix3()
    
        ' Reset all Solver settings
        SolverReset
    
        ' Check the value to determine Max or Min
        If Range("C19").Value = "Max" Then
            ' Execute the Solver
            SolverOk SetCell:="$B$13", MaxMinVal:=1, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C19").Select
        
        ElseIf Range("C19").Value = "Min" Then
            ' Execute the Solver
            SolverOk SetCell:="$B$13", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C19").Select
            
        Else
            MsgBox "The cell must contain 'Max' or 'Min'"
            Exit Sub
        End If
        
        ' Copy and paste
        Range("B13:L13").Copy
        Range("C35:M35").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("C35").Select
        Range("M4:M6").ClearContents
        
        ' Reset all Solver settings
        SolverReset

        ' Check the value to determine Max or Min
        If Range("C20").Value = "Max" Then
            ' Execute the Solver
            SolverOk SetCell:="$C$13", MaxMinVal:=1, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C20").Select
        
        ElseIf Range("C20").Value = "Min" Then
            ' Execute the Solver
            SolverOk SetCell:="$C$13", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C20").Select
            
        Else
            MsgBox "The cell must contain 'Max' or 'Min'"
            Exit Sub
        End If
        
        ' Copy and paste
        Range("B13:L13").Copy
        Range("C36:M36").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("D36").Select
        Range("M4:M6").ClearContents
        
        ' Reset all Solver settings
        SolverReset

        ' Check the value to determine Max or Min
        If Range("C21").Value = "Max" Then
            ' Execute the Solver
            SolverOk SetCell:="$D$13", MaxMinVal:=1, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C21").Select
        
        ElseIf Range("C21").Value = "Min" Then
            ' Execute the Solver
            SolverOk SetCell:="$D$13", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C21").Select
            
        Else
            MsgBox "The cell must contain 'Max' or 'Min'"
            Exit Sub
        End If
        
        ' Copy and paste
        Range("B13:L13").Copy
        Range("C37:M37").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("E37").Select
        Range("M4:M6").ClearContents
        
        ' Reset all Solver settings
        SolverReset
        
        ' Check the value to determine Max or Min
        If Range("C22").Value = "Max" Then
            ' Execute the Solver
            SolverOk SetCell:="$E$13", MaxMinVal:=1, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C22").Select
        
        ElseIf Range("C22").Value = "Min" Then
            ' Execute the Solver
            SolverOk SetCell:="$E$13", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C22").Select
            
        Else
            MsgBox "The cell must contain 'Max' or 'Min'"
            Exit Sub
        End If
        
        ' Copy and paste
        Range("B13:L13").Copy
        Range("C38:M38").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("F38").Select
        Range("M4:M6").ClearContents
        
        ' Reset all Solver settings
        SolverReset
        
        ' Check the value to determine Max or Min
        If Range("C23").Value = "Max" Then
            ' Execute the Solver
            SolverOk SetCell:="$F$13", MaxMinVal:=1, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C23").Select
        
        ElseIf Range("C23").Value = "Min" Then
            ' Execute the Solver
            SolverOk SetCell:="$F$13", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C23").Select
            
        Else
            MsgBox "The cell must contain 'Max' or 'Min'"
            Exit Sub
        End If
        
        ' Copy and paste
        Range("B13:L13").Copy
        Range("C39:M39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("G39").Select
        Range("M4:M6").ClearContents
        
        ' Reset all Solver settings
        SolverReset
        
        ' Check the value to determine Max or Min
        If Range("C24").Value = "Max" Then
            ' Execute the Solver
            SolverOk SetCell:="$G$13", MaxMinVal:=1, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C24").Select
        
        ElseIf Range("C24").Value = "Min" Then
            ' Execute the Solver
            SolverOk SetCell:="$G$13", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C24").Select
            
        Else
            MsgBox "The cell must contain 'Max' or 'Min'"
            Exit Sub
        End If
        
        ' Copy and paste
        Range("B13:L13").Copy
        Range("C40:M40").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("H40").Select
        Range("M4:M6").ClearContents
        
        ' Reset all Solver settings
        SolverReset
        
        ' Check the value to determine Max or Min
        If Range("C25").Value = "Max" Then
            ' Execute the Solver
            SolverOk SetCell:="$H$13", MaxMinVal:=1, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C25").Select
        
        ElseIf Range("C25").Value = "Min" Then
            ' Execute the Solver
            SolverOk SetCell:="$H$13", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C25").Select
            
        Else
            MsgBox "The cell must contain 'Max' or 'Min'"
            Exit Sub
        End If
        
        ' Copy and paste
        Range("B13:L13").Copy
        Range("C41:M41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("I41").Select
        Range("M4:M6").ClearContents
        
        ' Reset all Solver settings
        SolverReset
        
        ' Check the value to determine Max or Min
        If Range("C26").Value = "Max" Then
            ' Execute the Solver
            SolverOk SetCell:="$I$13", MaxMinVal:=1, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C26").Select
        
        ElseIf Range("C26").Value = "Min" Then
            ' Execute the Solver
            SolverOk SetCell:="$I$13", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C26").Select
            
        Else
            MsgBox "The cell must contain 'Max' or 'Min'"
            Exit Sub
        End If
        
        ' Copy and paste
        Range("B13:L13").Copy
        Range("C42:M42").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("J42").Select
        Range("M4:M6").ClearContents
        
        ' Reset all Solver settings
        SolverReset
        
        ' Check the value to determine Max or Min
        If Range("C27").Value = "Max" Then
            ' Execute the Solver
            SolverOk SetCell:="$J$13", MaxMinVal:=1, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C27").Select
        
        ElseIf Range("C27").Value = "Min" Then
            ' Execute the Solver
            SolverOk SetCell:="$J$13", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C27").Select
            
        Else
            MsgBox "The cell must contain 'Max' or 'Min'"
            Exit Sub
        End If
        
        ' Copy and paste
        Range("B13:L13").Copy
        Range("C43:M43").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("K43").Select
        Range("M4:M6").ClearContents
        
        ' Reset all Solver settings
        SolverReset
        
        ' Check the value to determine Max or Min
        If Range("C28").Value = "Max" Then
            ' Execute the Solver
            SolverOk SetCell:="$K$13", MaxMinVal:=1, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C28").Select
        
        ElseIf Range("C28").Value = "Min" Then
            ' Execute the Solver
            SolverOk SetCell:="$K$13", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C28").Select
            
        Else
            MsgBox "The cell must contain 'Max' or 'Min'"
            Exit Sub
        End If
        
        ' Copy and paste
        Range("B13:L13").Copy
        Range("C44:M44").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("L44").Select
        Range("M4:M6").ClearContents
        
        ' Reset all Solver settings
        SolverReset
        
        ' Check the value to determine Max or Min
        If Range("C29").Value = "Max" Then
            ' Execute the Solver
            SolverOk SetCell:="$L$13", MaxMinVal:=1, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C29").Select
        
        ElseIf Range("C29").Value = "Min" Then
            ' Execute the Solver
            SolverOk SetCell:="$L$13", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$M$4:$M$6").Offset(i, 0), Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
            0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
            SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
            :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
            IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
            SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$G$15"
            SolverOptions AssumeNonNeg:=False
            SolverSolve True
            Range("C29").Select
            
        Else
            MsgBox "The cell must contain 'Max' or 'Min'"
            Exit Sub
        End If
        
        ' Copy and paste
        Range("B13:L13").Copy
        Range("C45:M45").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("M45").Select
        Range("M4:M6").ClearContents
        
        ' Reset all Solver settings
        SolverReset
        

End Sub

'Calculate the Payoff matrix 8Y
Sub PayoffMatrix8()
    
    ' Reset all Solver settings
    SolverReset
    
    ' Check the value to determine Max or Min
    If Range("C19").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$B$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C19").Select
    
    ElseIf Range("C19").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$B$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C19").Select
        
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("C36:J36").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("C36").Select
    Range("J4:J6").ClearContents

    ' Reset all Solver settings
    SolverReset
    
    ' Check the value to determine Max or Min
    If Range("C20").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$C$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C20").Select
    
    ElseIf Range("C20").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$C$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C20").Select
        
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("C37:J37").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("D14").Select
    Range("J4:J6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C21").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$D$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C21").Select
        
    
    ElseIf Range("C21").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$D$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C21").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("C38:J38").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("E15").Select
    Range("J4:J6").ClearContents

    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C22").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$E$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C22").Select
        
    
    ElseIf Range("C22").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$E$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C22").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("C39:J39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("F16").Select
    Range("J4:J6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C23").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$F$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C23").Select
        
    
    ElseIf Range("C23").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$F$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C23").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("C40:J40").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("G17").Select
    Range("J4:J6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C24").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$G$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C24").Select
        
    
    ElseIf Range("C24").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$G$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C24").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("C41:J41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("H19").Select
    Range("J4:J6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C25").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$H$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C25").Select
        
    
    ElseIf Range("C25").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$H$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C25").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("C42:J42").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("I14").Select
    Range("J4:J6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

    ' Check the value to determine Max or Min
    If Range("C26").Value = "Max" Then
        ' Execute the Solver
        SolverOk SetCell:="$I$13", MaxMinVal:=1, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C26").Select
        
    
    ElseIf Range("C26").Value = "Min" Then
        ' Execute the Solver
        SolverOk SetCell:="$I$13", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverOptions MaxTime:=100, Iterations:=100, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, Derivatives:=1
        SolverOptions PopulationSize:=100, RandomSeed:=0, MutationRate:=0.075, Multistart _
        :=False, RequireBounds:=True, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=0.5, SolveWithout:=True, MaxTimeNoImp:=30
        SolverAdd CellRef:="$C$15", Relation:=1, FormulaText:="$F$19"
        SolverOptions AssumeNonNeg:=False
        SolverSolve True
        Range("C26").Select
    Else
        MsgBox "The cell must contain 'Max' or 'Min'"
        Exit Sub
    End If
    
    ' Copy and paste
    Range("B13:I13").Copy
    Range("C43:J43").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("J15").Select
    Range("J4:J6").ClearContents
    
    ' Reset all Solver settings
    SolverReset

End Sub

' Navigate to the points search tab (3Y)
Sub SearchPoints3()

    Range("DA3").Select

End Sub

' Navigate to the points search tab (8Y)
Sub SearchPoints8()

    Range("DI3").Select

End Sub

' Save the chosen points
Sub SavePoints3()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    
    ' Find the last filled row in column
    lastRow = ws.Cells(ws.Rows.Count, "DA").End(xlUp).Row
    
    ' If the last filled row is less, paste before
    If lastRow < 7 Then
        ws.Range("DA3:EB3").Copy
        ws.Range("DA7").PasteSpecial Paste:=xlPasteValues
    Else
        ' Otherwise, paste in the next empty row
        ws.Range("DA" & lastRow + 1 & ":EB" & lastRow + 1).Value = ws.Range("DA3:EB3").Value
    End If
    
    ' Apply validation for data entry only
    With ws.Range("DA" & lastRow + 1 & ":DB" & lastRow + 1).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' Remove borders from the new row
    With ws.Range("DA" & lastRow + 1 & ":EB" & lastRow + 1).Borders
        .LineStyle = xlNone
    End With
    
    ' Go back to initial cell
    ws.Range("DA7").Select
End Sub


' Save the chosen points
Sub SavePoints8()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    
    ' Find the last filled row in column
    lastRow = ws.Cells(ws.Rows.Count, "DI").End(xlUp).Row
    
    ' If the last filled row is less, paste before
    If lastRow < 7 Then
        ws.Range("DI3:EQ3").Copy
        ws.Range("DA7").PasteSpecial Paste:=xlPasteValues
    Else
        ' Otherwise, paste in the next empty rowa
        ws.Range("DI" & lastRow + 1 & ":EQ" & lastRow + 1).Value = ws.Range("DI3:EQ3").Value
    End If
    
    ' Apply validation for data entry only
    With ws.Range("DI" & lastRow + 1 & ":DJ" & lastRow + 1).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' Remove borders from the new row
    With ws.Range("DI" & lastRow + 1 & ":EQ" & lastRow + 1).Borders
        .LineStyle = xlNone
    End With
    
    ' Go back to initial cell
    ws.Range("DI7").Select
End Sub
