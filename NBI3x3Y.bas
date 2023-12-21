Attribute VB_Name = "NBI3x3Y"
'NBI 3x3Y Code
Private Sub Variables_Initialize()

    ' Add the options to the combo boxes
    ComboBox1.AddItem "Maximization"
    ComboBox1.AddItem "Minimization"
    ComboBox2.AddItem "Maximization"
    ComboBox2.AddItem "Minimization"
    ComboBox3.AddItem "Maximization"
    ComboBox3.AddItem "Minimization"
    
End Sub


Private Sub cmdsave_Click()
    
    'Set and store variables
    Dim TextVar1 As String
    Dim TextVar2 As String
    Dim TextVar3 As String
    
    Range("B2").Value = Var1.Value
    Range("C2").Value = Var2.Value
    Range("D2").Value = Var3.Value
    
    TextVar1 = Var1.Value
    TextVar2 = Var2.Value
    TextVar3 = Var3.Value

    If Var1.Value = "" Or Var2.Value = "" Or Var3.Value = "" Then
        MsgBox "Fill in all options to save.", vbExclamation, "Warning"
    ElseIf Not (VarMaxMin3.ComboBox1.Value = "Maximization" Or VarMaxMin3.ComboBox1.Value = "Minimization") Or Not (VarMaxMin3.ComboBox2.Value = "Maximization" Or VarMaxMin3.ComboBox2.Value = "Minimization") Or Not (VarMaxMin3.ComboBox3.Value = "Maximization" Or VarMaxMin3.ComboBox3.Value = "Minimization") Then
        MsgBox "Select the optimization direction to save.", vbExclamation, "Warning"
    Else
        If ComboBox1.Value = "Maximization" Then
            Range("C19").Value = "Maximization"
        ElseIf ComboBox1.Value = "Minimization" Then
            Range("C19").Value = "Minimization"
        End If
        
        If ComboBox2.Value = "Maximization" Then
            Range("C20").Value = "Maximization"
        ElseIf ComboBox2.Value = "Minimization" Then
            Range("C20").Value = "Minimization"
        End If
        If ComboBox3.Value = "Maximization" Then
            Range("C21").Value = "Maximization"
        ElseIf ComboBox3.Value = "Minimization" Then
            Range("C21").Value = "Minimization"
        End If
        
        Unload Me
        MsgBox ("The responses Y1, Y2, and Y3 were defined as " & TextVar1 & ", " & TextVar2 & " and " & TextVar3 & "!")
    End If

End Sub


Sub Variables()

    Application.ScreenUpdating = False
    
    Variables.Show
    
    Application.ScreenUpdating = True
    
End Sub


Sub IndOpti()

    'Reset all Solver settings
    SolverReset
    
    ' Check if the cells contain minimization or maximization and run the solver for individual optimization for the first variable
    If Range("C27").Value = "Maximization" Then
        ' Run the Solver
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
    
    ElseIf Range("C27").Value = "Minimization" Then
        ' Run the Solver
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
        MsgBox "The cell must contain 'Maximization' ou 'Minimization'"
        Exit Sub
    End If
    
    ' Copy and paste into the Payoff Matrix for the first variable
    Range("J13:L13").Copy
    Range("O3:O5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("J13").Select
    Range("M4:M6").ClearContents

    'Reset all Solver settings
    SolverReset
    
    ' Check if the cells contain minimization or maximization and run the solver for individual optimization for the second variable
    If Range("C28").Value = "Maximization" Then
        ' Run the Solver
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
    
    ElseIf Range("C28").Value = "Minimization" Then
        ' Run the Solver
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
        MsgBox "The cell must contain 'Maximization' ou 'Minimization'"
        Exit Sub
    End If
    
    ' Copy and paste into the Payoff Matrix for the second variable
    Range("J13:L13").Copy
    Range("P3:P5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("K13").Select
    Range("M4:M6").ClearContents
    
    'Reset all Solver settings
    SolverReset

    ' Check if the cells contain minimization or maximization and run the solver for individual optimization for the third variable
    If Range("C29").Value = "Maximization" Then
        ' Run the Solver
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
        
    
    ElseIf Range("C29").Value = "Minimization" Then
        ' Run the Solver
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
        MsgBox "The cell must contain 'Maximization' ou 'Minimization'"
        Exit Sub
    End If
    
    ' Copy and paste into the Payoff Matrix for the third variable
    Range("J13:L13").Copy
    Range("Q3:Q5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("L13").Select
    Range("M4:M6").ClearContents
    
    Application.CutCopyMode = False
    Range("A1").Select
    
    'Check and set the Nadir and Utopia formula according to the optimization direction
    If Range("C27").Value = "Maximization" Then
        Range("S3").FormulaLocal = "=MAX(O3:Q3)"
        Range("T3").FormulaLocal = "=MIN(O3:Q3)"
        Range("AU3").FormulaLocal = "=MAX(AK3:AK68)"
    ElseIf Range("C27").Value = "Minimization" Then
        Range("S3").FormulaLocal = "=MIN(O3:Q3)"
        Range("T3").FormulaLocal = "=MAX(O3:Q3)"
        Range("AU3").FormulaLocal = "=MIN(AK3:AK68)"
    Else
        MsgBox "The cell must contain 'Maximization' ou 'Minimization'"
    End If
    
    If Range("C28").Value = "Maximization" Then
        Range("S4").FormulaLocal = "=MAX(O4:Q4)"
        Range("T4").FormulaLocal = "=MIN(O4:Q4)"
        Range("AU4").FormulaLocal = "=MAX(AL3:AL68)"
    ElseIf Range("C28").Value = "Minimization" Then
        Range("S4").FormulaLocal = "=MIN(O4:Q4)"
        Range("T4").FormulaLocal = "=MAX(O4:Q4)"
        Range("AU4").FormulaLocal = "=MIN(AL3:AL68)"
    Else
        MsgBox "The cell must contain 'Maximization' ou 'Minimization'"
    End If
    
    If Range("C29").Value = "Maximization" Then
        Range("S5").FormulaLocal = "=MAX(O5:Q5)"
        Range("T5").FormulaLocal = "=MIN(O5:Q5)"
        Range("AU5").FormulaLocal = "=MAX(AM3:AM68)"
    ElseIf Range("C29").Value = "Minimization" Then
        Range("S5").FormulaLocal = "=MIN(O5:Q5)"
        Range("T5").FormulaLocal = "=MAX(O5:Q5)"
        Range("AU5").FormulaLocal = "=MIN(AM3:AM68)"
    Else
        MsgBox "The cell must contain 'Maximization' OR 'Minimization'"
    End If

    ' Replace "@" with "=" in the formulas
    Range("AU3:AU5").Replace What:="@", Replacement:="", LookAt:=xlPart
    
    SolverReset

End Sub


Sub NBI3Solve()

    ' Iterations with the loop from 1 to 66
    For i = 1 To 66
        ' Set the value of i
        Range("$T$14").Value = i
        
        'Run the Solver
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
        
        ' Copy and paste the results to the table with all iterations
        Range("$C$32:$P$32").Copy
        Range("Z" & (i + 2) & ":AM" & (i + 2)).PasteSpecial Paste:=xlPasteValues
  
        Range("C15").Copy
        Range("AW" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("E15:E17").Copy
        Range("AX" & (i + 2) & ":AZ" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Application.CutCopyMode = False
        
        'Reset all Solver settings
        SolverReset
        
    Next i
    
    Range("T14").Value = 1
    Range("T14").Select
    Range("M4:M6").ClearContents
    
    'Fix Borders
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
        Range("AW3:AZ68").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AW3:AZ68").Select
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
    
    MsgBox "NBI completed!"
    Range("A1").Select

End Sub
