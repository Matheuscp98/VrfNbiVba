Attribute VB_Name = "NBI_Completed"
Sub LigarTelaCheia()

    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    
End Sub

Sub DesligarTelaCheia()

    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    
End Sub

Sub SalvarPlanilha()

    Application.ScreenUpdating = False

    ThisWorkbook.Save

    Application.ScreenUpdating = True

End Sub

Sub Guia2Y()
    Sheets("NBI 2Y").Activate
    Range("A1").Select
End Sub

Sub Guia3Y()
    Sheets("NBI 3Y").Activate
    Range("A1").Select
End Sub

Sub Guia4Y()
    Sheets("NBI 4").Activate
    Range("A1").Select
End Sub


Sub LimparCelulas3()

    ' Limpar o intervalo resultados Solver
    Range("Z3:AM68").ClearContents
    Range("Z74:AM139").ClearContents
    Range("Z145:AM210").ClearContents
    
    ' Limpar o intervalo x1 até x3
    Range("M4:M6").ClearContents
    
    ' Limpar o intervalo de restrições
    Range("AX3:BA68").ClearContents
    Range("AX74:BA139").ClearContents
    Range("AX145:BA210").ClearContents
    
    'Voltar o n para o valor inicial
    Range("T14").Value = 1
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Retornar à célula A1
    Range("A1").Select
    
End Sub
Sub LimparCelulas8()

    ' Limpar o intervalo resultados Solver
    Range("AG3:AQ794").ClearContents
    Range("AG799:AQ1590").ClearContents
    Range("AG1595:AQ2386").ClearContents
    
    ' Limpar o intervalo x1 até x3
    Range("J4:J6").ClearContents
    
    ' Limpar o intervalo de restrições
    Range("BA3:BI794").ClearContents
    Range("BA799:BI1590").ClearContents
    Range("BA1595:BI2386").ClearContents
    
    'Voltar o n para o valor inicial
    Range("V23").Value = 1
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Retornar à célula A1
    Range("A1").Select
    
End Sub

Sub LimparCelulasPost()

    ' Limpar o intervalo resultados Solver
    Range("M3:W23").ClearContents
    
    ' Limpar o intervalo x1 até x3
    Range("G3:G5").ClearContents
    
    'Voltar o n e BetaMD para o valor inicial
    Range("E27").Value = 1
    Range("C33").Value = 1
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Retornar à célula A1
    Range("A1").Select
    
    
End Sub

Sub LimparPontos()

    ' Limpar o intervalo dos pontos
    Range("DA7:EB10000").ClearContents
    Range("DA3").Select

End Sub

Sub LimparPontos8()

    ' Limpar o intervalo dos pontos
    Range("DI7:EQ10000").ClearContents
    Range("DI3").Select

End Sub

Sub NBIASolve()

    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Loop de 1 a 66 para T14
    For i = 1 To 66
        ' Defina o valor de T14
        Range("$T$14").Value = i
        
        
        'Executar o Solver
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
        
        ' Copie e cole os resultados
        Range("$C$32:$P$32").Copy
        Range("Z" & (i + 73) & ":AM" & (i + 73)).PasteSpecial Paste:=xlPasteValues
        
        Range("C15").Copy
        Range("AX" & (i + 73)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("E15:E17").Copy
        Range("AY" & (i + 73) & ":BA" & (i + 73)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Application.CutCopyMode = False
        
        'Resetar todas as configurações do Solver
        SolverReset
        
    Next i
    
    
    Range("T14").Value = 1
    Range("T14").Select
    Range("M4:M6").ClearContents
    
    'Corrigir Bordas
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
    
    MsgBox "NBI concluído com os pontos anteriores a cada iteração mantidos!"
    Range("A1").Select

End Sub

Sub NBIA8Solve()

    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Loop de 1 a 792 para V23
    For i = 1 To 792
        ' Defina o valor de V23
        Range("$V$23").Value = i
        
        'Executar o Solver
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
        
        ' Copie e cole os resultados
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
        
        'Resetar todas as configurações do Solver
        SolverReset
        
    Next i
    
    
    Range("V23").Value = 1
    Range("V23").Select
    Range("J4:J6").ClearContents
    
    'Corrigir Bordas
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
    
    MsgBox "NBI concluído com os pontos anteriores a cada iteração mantidos!"
    Range("A1").Select

End Sub

Sub NBIAPost()

    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Loop de 1 a 21 para B31
    For i = 1 To 21
    ' Defina o valor de B31
    Range("$C$33").Value = i
    
        ' Executar o Solver
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
        
        ' Copie e cole
        Range("$B$31:$J$31").Copy
        Range("L" & (i + 2) & ":T" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("C20").Copy
        Range("V" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("E20").Copy
        Range("W" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Application.CutCopyMode = False


        'Resetar todas as configurações do Solver
        SolverReset
        
        'Alterar Beta
        Range("E27").Value = Range("E27").Value - 0.05
        
    Next i
    
    Range("E27").Value = 1
    Range("C33").Value = 1
    Range("G3:G5").ClearContents
    
    MsgBox "NBI concluído com os pontos anteriores a cada iteração mantidos!"
    Range("A1").Select


End Sub


Sub NBIZPost()

    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Loop de 1 a 21 para B31
    For i = 1 To 21
    ' Defina o valor de B31
    Range("$C$33").Value = i
    
    Range("G3").Value = 0
    Range("G4").Value = 0
    Range("G5").Value = 0
    
        ' Executar o Solver
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
        
        ' Copie e cole
        Range("$B$31:$J$31").Copy
        Range("L" & (i + 2) & ":T" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("C20").Copy
        Range("V" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("E20").Copy
        Range("W" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Application.CutCopyMode = False


        'Resetar todas as configurações do Solver
        SolverReset
        
        'Alterar Beta
        Range("E27").Value = Range("E27").Value - 0.05
        
    Next i
    
    Range("E27").Value = 1
    Range("C33").Value = 1
    Range("G3:G5").ClearContents
    
    MsgBox "NBI concluído com os pontos anteriores a cada iteração mantidos!"
    Range("A1").Select


End Sub


Sub NBIOSolve()

    'Resetar todas as configurações do Solver
    SolverReset

    ' Loop de 1 a 66 para T14
    For i = 1 To 66
        ' Defina o valor de T14
        Range("$T$14").Value = i
        
        Range("M4").Value = Range("Z69").Value
        Range("M5").Value = Range("AA69").Value
        Range("M6").Value = Range("AB69").Value
        
        'Executar o Solver
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
        
        ' Copie e cole os resultados
        Range("$C$32:$P$32").Copy
        Range("Z" & (i + 144) & ":AM" & (i + 144)).PasteSpecial Paste:=xlPasteValues
        Range("C15").Copy
        Range("AX" & (i + 144)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("E15:E17").Copy
        Range("AY" & (i + 144) & ":BA" & (i + 144)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Application.CutCopyMode = False
        
        'Resetar todas as configurações do Solver
        SolverReset
        
    Next i
    
    
    Range("T14").Value = 1
    Range("T14").Select
    Range("M4:M6").ClearContents
    
    'Corrigir Bordas
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
    
    MsgBox "NBI concluído com os pontos ótimos dos fatores rotacionados mantidos a cada iteração!"
    Range("A1").Select

End Sub

Sub NBIO8Solve()

    'Resetar todas as configurações do Solver
    SolverReset

    ' Loop de 1 a 792 para V23
    For i = 1 To 792
        ' Defina o valor de V23
        Range("$V$23").Value = i
        
        Range("J4").Value = Range("AG797").Value
        Range("J5").Value = Range("AH797").Value
        Range("J6").Value = Range("AI797").Value
        
        'Executar o Solver
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
        
        ' Copie e cole os resultados
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
        
        'Resetar todas as configurações do Solver
        SolverReset
        
    Next i
    
    
    Range("V23").Value = 1
    Range("V23").Select
    Range("J4:J6").ClearContents
    
    'Corrigir Bordas
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
    
    MsgBox "NBI concluído com os pontos ótimos dos fatores rotacionados mantidos a cada iteração!"
    Range("A1").Select

End Sub



Sub NBIPostRSM()

    'Executar o Solver
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
        
    ' Copie e cole os resultados
    Range("$B$13:$L$13").Copy
    Range("$C$38:$M$38").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False

    'Resetar todas as configurações do Solver
    SolverReset

    Range("U22").Select
    
End Sub

Sub NBIZSolve()

    'Resetar todas as configurações do Solver
    SolverReset

    ' Loop de 1 a 66 para T14
    For i = 1 To 66
        ' Defina o valor de T14
        Range("$T$14").Value = i
        
        Range("M4").Value = 0
        Range("M5").Value = 0
        Range("M6").Value = 0
        
        'Executar o Solver
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
        
        ' Copie e cole os resultados
        Range("$C$32:$P$32").Copy
        Range("Z" & (i + 2) & ":AM" & (i + 2)).PasteSpecial Paste:=xlPasteValues
        Range("C15").Copy
        Range("AX" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("E15:E17").Copy
        Range("AY" & (i + 2) & ":BA" & (i + 2)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Application.CutCopyMode = False
        
        'Resetar todas as configurações do Solver
        SolverReset
        
    Next i
    
    
    Range("T14").Value = 1
    Range("T14").Select
    Range("M4:M6").ClearContents
    
    'Corrigir Bordas
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
    
    MsgBox "NBI concluído com os pontos anteriores a cada iteração zerados!"
    Range("A1").Select

End Sub

Sub NBIZ8Solve()

    'Resetar todas as configurações do Solver
    SolverReset

    ' Loop de 1 a 792 para V23
    For i = 1 To 792
        ' Defina o valor de V23
        Range("$V$23").Value = i
        
        Range("J4").Value = 0
        Range("J5").Value = 0
        Range("J6").Value = 0
        
        'Executar o Solver
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
        
        ' Copie e cole os resultados
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
        
        'Resetar todas as configurações do Solver
        SolverReset
        
    Next i
    
    
    Range("V23").Value = 1
    Range("V23").Select
    Range("J4:J6").ClearContents
    
    'Corrigir Bordas
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
    
    MsgBox "NBI concluído com os pontos anteriores a cada iteração zerados!"
    Range("A1").Select

End Sub



Sub OtiInd3()

    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Verifique o valor para determinar Max ou Min
    If Range("C27").Value = "Maximização" Then
        ' Executar o Solver
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
    
    ElseIf Range("C27").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("J13:L13").Copy
    Range("O3:O5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("J13").Select
    Range("M4:M6").ClearContents

    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Verifique o valor para determinar Max ou Min
    If Range("C28").Value = "Maximização" Then
        ' Executar o Solver
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
    
    ElseIf Range("C28").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("J13:L13").Copy
    Range("P3:P5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("K13").Select
    Range("M4:M6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C29").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C29").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("J13:L13").Copy
    Range("Q3:Q5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("L13").Select
    Range("M4:M6").ClearContents
    
    Application.CutCopyMode = False
    Range("A1").Select
    
    ' Verifique C27 e defina a fórmula em S3, T3 e AU3
    If Range("C27").Value = "Maximização" Then
        Range("S3").FormulaLocal = "=MÁXIMO(O3:Q3)"
        Range("T3").FormulaLocal = "=MÍNIMO(O3:Q3)"
        Range("AV3").FormulaLocal = "=MÁXIMO(AK3:AK68)"
    ElseIf Range("C27").Value = "Minimização" Then
        Range("S3").FormulaLocal = "=MÍNIMO(O3:Q3)"
        Range("T3").FormulaLocal = "=MÁXIMO(O3:Q3)"
        Range("AV3").FormulaLocal = "=MÍNIMO(AK3:AK68)"
    Else
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
    End If
    
    ' Verifique C28 e defina a fórmula em S4, T4 e AU4
    If Range("C28").Value = "Maximização" Then
        Range("S4").FormulaLocal = "=MÁXIMO(O4:Q4)"
        Range("T4").FormulaLocal = "=MÍNIMO(O4:Q4)"
        Range("AV4").FormulaLocal = "=MÁXIMO(AL3:AL68)"
    ElseIf Range("C28").Value = "Minimização" Then
        Range("S4").FormulaLocal = "=MÍNIMO(O4:Q4)"
        Range("T4").FormulaLocal = "=MÁXIMO(O4:Q4)"
        Range("AV4").FormulaLocal = "=MÍNIMO(AL3:AL68)"
    Else
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
    End If
    
    ' Verifique C29 e defina a fórmula em S5, T5 e AU5
    If Range("C29").Value = "Maximização" Then
        Range("S5").FormulaLocal = "=MÁXIMO(O5:Q5)"
        Range("T5").FormulaLocal = "=MÍNIMO(O5:Q5)"
        Range("AV5").FormulaLocal = "=MÁXIMO(AM3:AM68)"
    ElseIf Range("C29").Value = "Minimização" Then
        Range("S5").FormulaLocal = "=MÍNIMO(O5:Q5)"
        Range("T5").FormulaLocal = "=MÁXIMO(O5:Q5)"
        Range("AV5").FormulaLocal = "=MÍNIMO(AM3:AM68)"
    Else
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
    End If

    ' Substituir "@" por "=" nas fórmulas
    Range("AV3:AV5").Replace What:="@", Replacement:="", LookAt:=xlPart
    
    'Resetar todas as configurações do Solver
    SolverReset

End Sub

Sub OtiInd8()

    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Verifique o valor para determinar Max ou Min
    If Range("C19").Value = "Maximização" Then
        ' Executar o Solver
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
    
    ElseIf Range("C19").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("L3:L10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("B13").Select
    Range("J4:J6").ClearContents

    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Verifique o valor para determinar Max ou Min
    If Range("C20").Value = "Maximização" Then
        ' Executar o Solver
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
    
    ElseIf Range("C20").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("M3:M10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("C13").Select
    Range("J4:J6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C21").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C21").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("N3:N10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("D13").Select
    Range("J4:J6").ClearContents

    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C22").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C22").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("O3:O10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("E13").Select
    Range("J4:J6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C23").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C23").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("P3:P10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("F13").Select
    Range("J4:J6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C24").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C24").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("Q3:Q10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("G13").Select
    Range("J4:J6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C25").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C25").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("R3:R10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("H13").Select
    Range("J4:J6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C26").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C26").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("S3:S10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("I13").Select
    Range("J4:J6").ClearContents
    

    Application.CutCopyMode = False
    Range("A1").Select
    
    ' Verifique C19 e defina a fórmula em U3, V3 e AY3
    If Range("C19").Value = "Maximização" Then
        Range("U3").FormulaLocal = "=MÁXIMO(L3:S3)"
        Range("V3").FormulaLocal = "=MÍNIMO(L3:S3)"
        Range("AY3").FormulaLocal = "=MÁXIMO(AJ3:AJ794)"
    ElseIf Range("C19").Value = "Minimização" Then
        Range("U3").FormulaLocal = "=MÍNIMO(L3:S3)"
        Range("V3").FormulaLocal = "=MÁXIMO(L3:S3)"
        Range("AY3").FormulaLocal = "=MÍNIMO(AJ3:AJ794)"
    Else
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
    End If
    
    ' Verifique C20 e defina a fórmula em U4, V4 e AY4
    If Range("C20").Value = "Maximização" Then
        Range("U4").FormulaLocal = "=MÁXIMO(L4:S4)"
        Range("V4").FormulaLocal = "=MÍNIMO(L4:S4)"
        Range("AY4").FormulaLocal = "=MÁXIMO(AK3:AK794)"
    ElseIf Range("C20").Value = "Minimização" Then
        Range("U4").FormulaLocal = "=MÍNIMO(L4:S4)"
        Range("V4").FormulaLocal = "=MÁXIMO(L4:S4)"
        Range("AY4").FormulaLocal = "=MÍNIMO(AK3:AK794)"
    Else
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
    End If
    
    ' Verifique C21 e defina a fórmula em U5, V5 e AY5
    If Range("C21").Value = "Maximização" Then
        Range("U3").FormulaLocal = "=MÁXIMO(L5:S5)"
        Range("V5").FormulaLocal = "=MÍNIMO(L5:S5)"
        Range("AY5").FormulaLocal = "=MÁXIMO(AL3:AL794)"
    ElseIf Range("C21").Value = "Minimização" Then
        Range("U5").FormulaLocal = "=MÍNIMO(L5:S5)"
        Range("V5").FormulaLocal = "=MÁXIMO(L5:S5)"
        Range("AY5").FormulaLocal = "=MÍNIMO(AL3:AL794)"
    Else
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
    End If
    
    ' Verifique C22 e defina a fórmula em U6, V6 e AY6
    If Range("C22").Value = "Maximização" Then
        Range("U6").FormulaLocal = "=MÁXIMO(L6:S6)"
        Range("V6").FormulaLocal = "=MÍNIMO(L6:S6)"
        Range("AY6").FormulaLocal = "=MÁXIMO(AM3:AM794)"
    ElseIf Range("C22").Value = "Minimização" Then
        Range("U6").FormulaLocal = "=MÍNIMO(L6:S6)"
        Range("V6").FormulaLocal = "=MÁXIMO(L6:S6)"
        Range("AY6").FormulaLocal = "=MÍNIMO(AM3:AM794)"
    Else
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
    End If
    
    ' Verifique C23 e defina a fórmula em U7, V7 e AY7
    If Range("C23").Value = "Maximização" Then
        Range("U7").FormulaLocal = "=MÁXIMO(L7:S7)"
        Range("V7").FormulaLocal = "=MÍNIMO(L7:S7)"
        Range("AY7").FormulaLocal = "=MÁXIMO(AN3:AN794)"
    ElseIf Range("C23").Value = "Minimização" Then
        Range("U7").FormulaLocal = "=MÍNIMO(L7:S7)"
        Range("V7").FormulaLocal = "=MÁXIMO(L7:S7)"
        Range("AY7").FormulaLocal = "=MÍNIMO(AN3:AN794)"
    Else
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
    End If
    
    ' Verifique C24 e defina a fórmula em U8, V8 e AY8
    If Range("C24").Value = "Maximização" Then
        Range("U8").FormulaLocal = "=MÁXIMO(L8:S8)"
        Range("V8").FormulaLocal = "=MÍNIMO(L8:S8)"
        Range("AY8").FormulaLocal = "=MÁXIMO(AO3:AO794)"
    ElseIf Range("C24").Value = "Minimização" Then
        Range("U8").FormulaLocal = "=MÍNIMO(L8:S8)"
        Range("V8").FormulaLocal = "=MÁXIMO(L8:S8)"
        Range("AY8").FormulaLocal = "=MÍNIMO(AO3:AO794)"
    Else
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
    End If
    
    ' Verifique C25 e defina a fórmula em U9, V9 e AY9
    If Range("C25").Value = "Maximização" Then
        Range("U9").FormulaLocal = "=MÁXIMO(L9:S9)"
        Range("V9").FormulaLocal = "=MÍNIMO(L9:S9)"
        Range("AY9").FormulaLocal = "=MÁXIMO(AP3:AP794)"
    ElseIf Range("C25").Value = "Minimização" Then
        Range("U9").FormulaLocal = "=MÍNIMO(L9:S9)"
        Range("V9").FormulaLocal = "=MÁXIMO(L9:S9)"
        Range("AY9").FormulaLocal = "=MÍNIMO(AP3:AP794)"
    Else
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
    End If
    
    ' Verifique C26 e defina a fórmula em U10, V10 e AY10
    If Range("C26").Value = "Maximização" Then
        Range("U10").FormulaLocal = "=MÁXIMO(L10:S10)"
        Range("V10").FormulaLocal = "=MÍNIMO(L10:S10)"
        Range("AY10").FormulaLocal = "=MÁXIMO(AQ3:AQ794)"
    ElseIf Range("C26").Value = "Minimização" Then
        Range("U10").FormulaLocal = "=MÍNIMO(L10:S10)"
        Range("V10").FormulaLocal = "=MÁXIMO(L10:S10)"
        Range("AY10").FormulaLocal = "=MÍNIMO(AQ3:AQ794)"
    Else
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
    End If
    
    ' Substituir "@" por "=" nas fórmulas
    Range("AY3:AY10").Replace What:="@", Replacement:="", LookAt:=xlPart
    
    'Resetar todas as configurações do Solver
    SolverReset

End Sub



Sub OtiIndPost()

    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Executar o Solver
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
    
    ' Copie e cole
    Range("C18:D18").Copy
    Range("I3:I4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("E18:F18").Copy
    Range("I7:I8").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("G3:G5").ClearContents

    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Executar o Solver
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
    
    ' Copie e cole
    Range("C18:D18").Copy
    Range("J3:J4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("E18:F18").Copy
    Range("J7:J8").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("G3:G5").ClearContents

    SolverReset
    
    Range("A1").Select
    
    
End Sub

Sub TabOti()
    
        'Resetar todas as configurações do Solver
        SolverReset
    
        ' Verifique o valor para determinar Max ou Min
        If Range("C19").Value = "Maximização" Then
            ' Executar o Solver
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
        
        ElseIf Range("C19").Value = "Minimização" Then
            ' Executar o Solver
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
            MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
            Exit Sub
        End If
        
        ' Copie e cole
        Range("B13:L13").Copy
        Range("C35:M35").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("C35").Select
        Range("M4:M6").ClearContents
        
        'Resetar todas as configurações do Solver
        SolverReset

        ' Verifique o valor para determinar Max ou Min
        If Range("C20").Value = "Maximização" Then
            ' Executar o Solver
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
        
        ElseIf Range("C20").Value = "Minimização" Then
            ' Executar o Solver
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
            MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
            Exit Sub
        End If
        
        ' Copie e cole
        Range("B13:L13").Copy
        Range("C36:M36").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("D36").Select
        Range("M4:M6").ClearContents
        
        'Resetar todas as configurações do Solver
        SolverReset

        ' Verifique o valor para determinar Max ou Min
        If Range("C21").Value = "Maximização" Then
            ' Executar o Solver
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
        
        ElseIf Range("C21").Value = "Minimização" Then
            ' Executar o Solver
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
            MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
            Exit Sub
        End If
        
        ' Copie e cole
        Range("B13:L13").Copy
        Range("C37:M37").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("E37").Select
        Range("M4:M6").ClearContents
        
        'Resetar todas as configurações do Solver
        SolverReset
        
        ' Verifique o valor para determinar Max ou Min
        If Range("C22").Value = "Maximização" Then
            ' Executar o Solver
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
        
        ElseIf Range("C22").Value = "Minimização" Then
            ' Executar o Solver
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
            MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
            Exit Sub
        End If
        
        ' Copie e cole
        Range("B13:L13").Copy
        Range("C38:M38").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("F38").Select
        Range("M4:M6").ClearContents
        
        'Resetar todas as configurações do Solver
        SolverReset
        
        ' Verifique o valor para determinar Max ou Min
        If Range("C23").Value = "Maximização" Then
            ' Executar o Solver
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
        
        ElseIf Range("C23").Value = "Minimização" Then
            ' Executar o Solver
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
            MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
            Exit Sub
        End If
        
        ' Copie e cole
        Range("B13:L13").Copy
        Range("C39:M39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("G39").Select
        Range("M4:M6").ClearContents
        
        'Resetar todas as configurações do Solver
        SolverReset
        
        ' Verifique o valor para determinar Max ou Min
        If Range("C24").Value = "Maximização" Then
            ' Executar o Solver
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
        
        ElseIf Range("C24").Value = "Minimização" Then
            ' Executar o Solver
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
            MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
            Exit Sub
        End If
        
        ' Copie e cole
        Range("B13:L13").Copy
        Range("C40:M40").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("H40").Select
        Range("M4:M6").ClearContents
        
        'Resetar todas as configurações do Solver
        SolverReset
        
        ' Verifique o valor para determinar Max ou Min
        If Range("C25").Value = "Maximização" Then
            ' Executar o Solver
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
        
        ElseIf Range("C25").Value = "Minimização" Then
            ' Executar o Solver
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
            MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
            Exit Sub
        End If
        
        ' Copie e cole
        Range("B13:L13").Copy
        Range("C41:M41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("I41").Select
        Range("M4:M6").ClearContents
        
        'Resetar todas as configurações do Solver
        SolverReset
        
        ' Verifique o valor para determinar Max ou Min
        If Range("C26").Value = "Maximização" Then
            ' Executar o Solver
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
        
        ElseIf Range("C26").Value = "Minimização" Then
            ' Executar o Solver
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
            MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
            Exit Sub
        End If
        
        ' Copie e cole
        Range("B13:L13").Copy
        Range("C42:M42").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("J42").Select
        Range("M4:M6").ClearContents
        
        'Resetar todas as configurações do Solver
        SolverReset
        
        ' Verifique o valor para determinar Max ou Min
        If Range("C27").Value = "Maximização" Then
            ' Executar o Solver
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
        
        ElseIf Range("C27").Value = "Minimização" Then
            ' Executar o Solver
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
            MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
            Exit Sub
        End If
        
        ' Copie e cole
        Range("B13:L13").Copy
        Range("C43:M43").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("K43").Select
        Range("M4:M6").ClearContents
        
        'Resetar todas as configurações do Solver
        SolverReset
        
        ' Verifique o valor para determinar Max ou Min
        If Range("C28").Value = "Maximização" Then
            ' Executar o Solver
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
        
        ElseIf Range("C28").Value = "Minimização" Then
            ' Executar o Solver
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
            MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
            Exit Sub
        End If
        
        ' Copie e cole
        Range("B13:L13").Copy
        Range("C44:M44").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("L44").Select
        Range("M4:M6").ClearContents
        
        'Resetar todas as configurações do Solver
        SolverReset
        
        ' Verifique o valor para determinar Max ou Min
        If Range("C29").Value = "Maximização" Then
            ' Executar o Solver
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
        
        ElseIf Range("C29").Value = "Minimização" Then
            ' Executar o Solver
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
            MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
            Exit Sub
        End If
        
        ' Copie e cole
        Range("B13:L13").Copy
        Range("C45:M45").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        Range("M45").Select
        Range("M4:M6").ClearContents
        
        'Resetar todas as configurações do Solver
        SolverReset
        

End Sub


Sub TabOti8()
    
    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Verifique o valor para determinar Max ou Min
    If Range("C19").Value = "Maximização" Then
        ' Executar o Solver
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
    
    ElseIf Range("C19").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("C36:J36").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("C36").Select
    Range("J4:J6").ClearContents

    'Resetar todas as configurações do Solver
    SolverReset
    
    ' Verifique o valor para determinar Max ou Min
    If Range("C20").Value = "Maximização" Then
        ' Executar o Solver
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
    
    ElseIf Range("C20").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("C37:J37").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("D14").Select
    Range("J4:J6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C21").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C21").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("C38:J38").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("E15").Select
    Range("J4:J6").ClearContents

    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C22").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C22").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("C39:J39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("F16").Select
    Range("J4:J6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C23").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C23").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("C40:J40").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("G17").Select
    Range("J4:J6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C24").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C24").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("C41:J41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("H19").Select
    Range("J4:J6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C25").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C25").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("C42:J42").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("I14").Select
    Range("J4:J6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

    ' Verifique o valor para determinar Max ou Min
    If Range("C26").Value = "Maximização" Then
        ' Executar o Solver
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
        
    
    ElseIf Range("C26").Value = "Minimização" Then
        ' Executar o Solver
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
        MsgBox "A célula deve conter 'Maximização' ou 'Minimização'"
        Exit Sub
    End If
    
    ' Copie e cole
    Range("B13:I13").Copy
    Range("C43:J43").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
    Range("J15").Select
    Range("J4:J6").ClearContents
    
    'Resetar todas as configurações do Solver
    SolverReset

End Sub

Sub PesqPontos()

    Range("DA3").Select

End Sub

Sub PesqPontos8()

    Range("DI3").Select

End Sub


Sub SalvarPontos()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    
    ' Encontra a última linha preenchida na coluna DA
    lastRow = ws.Cells(ws.Rows.Count, "DA").End(xlUp).Row
    
    ' Se a última linha preenchida for menor que 7, colar em DA7
    If lastRow < 7 Then
        ws.Range("DA3:EB3").Copy
        ws.Range("DA7").PasteSpecial Paste:=xlPasteValues
    Else
        ' Caso contrário, colar na próxima linha vazia
        ws.Range("DA" & lastRow + 1 & ":EB" & lastRow + 1).Value = ws.Range("DA3:EB3").Value
    End If
    
    ' Aplica validação apenas para entrada de dados
    With ws.Range("DA" & lastRow + 1 & ":DB" & lastRow + 1).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' Remove as bordas da nova linha
    With ws.Range("DA" & lastRow + 1 & ":EB" & lastRow + 1).Borders
        .LineStyle = xlNone
    End With
    
    ' Volta para a célula DA7
    ws.Range("DA7").Select
End Sub


Sub SalvarPontos8()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    
    ' Encontra a última linha preenchida na coluna DA
    lastRow = ws.Cells(ws.Rows.Count, "DI").End(xlUp).Row
    
    ' Se a última linha preenchida for menor que 7, colar em DA7
    If lastRow < 7 Then
        ws.Range("DI3:EQ3").Copy
        ws.Range("DA7").PasteSpecial Paste:=xlPasteValues
    Else
        ' Caso contrário, colar na próxima linha vazia
        ws.Range("DI" & lastRow + 1 & ":EQ" & lastRow + 1).Value = ws.Range("DI3:EQ3").Value
    End If
    
    ' Aplica validação apenas para entrada de dados
    With ws.Range("DI" & lastRow + 1 & ":DJ" & lastRow + 1).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' Remove as bordas da nova linha
    With ws.Range("DI" & lastRow + 1 & ":EQ" & lastRow + 1).Borders
        .LineStyle = xlNone
    End With
    
    ' Volta para a célula DA7
    ws.Range("DI7").Select
End Sub



Sub Var3MaxMin()

    Application.ScreenUpdating = False
    
    VarMaxMin3.Show
    
    Application.ScreenUpdating = True
    
End Sub



Sub Var8MaxMin()

    Application.ScreenUpdating = False
    
    VarMaxMin8.Show
    
    Application.ScreenUpdating = True
    
End Sub


Sub Volt()

    Range("A1").Select

End Sub

Private Sub UserForm_Initialize()

    ' Adicionar as opções às caixas de combinação
    ComboBox1.AddItem "Maximização"
    ComboBox1.AddItem "Minimização"
    
    ComboBox2.AddItem "Maximização"
    ComboBox2.AddItem "Minimização"
    
    ComboBox3.AddItem "Maximização"
    ComboBox3.AddItem "Minimização"
    
End Sub

Private Sub cmdgravar1_Click()
    
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
        MsgBox "Preencha todas as opções para salvar.", vbExclamation, "Aviso"
    ElseIf Not (VarMaxMin3.ComboBox1.Value = "Maximização" Or VarMaxMin3.ComboBox1.Value = "Minimização") Or Not (VarMaxMin3.ComboBox2.Value = "Maximização" Or VarMaxMin3.ComboBox2.Value = "Minimização") Or Not (VarMaxMin3.ComboBox3.Value = "Maximização" Or VarMaxMin3.ComboBox3.Value = "Minimização") Then
        MsgBox "Selecione o sentido da otimização para salvar.", vbExclamation, "Aviso"
    Else
        If ComboBox1.Value = "Maximização" Then
            Range("C19").Value = "Maximização"
        ElseIf ComboBox1.Value = "Minimização" Then
            Range("C19").Value = "Minimização"
        End If
        
        If ComboBox2.Value = "Maximização" Then
            Range("C20").Value = "Maximização"
        ElseIf ComboBox2.Value = "Minimização" Then
            Range("C20").Value = "Minimização"
        End If
        If ComboBox3.Value = "Maximização" Then
            Range("C21").Value = "Maximização"
        ElseIf ComboBox3.Value = "Minimização" Then
            Range("C21").Value = "Minimização"
        End If
        
        Unload Me
        MsgBox ("As respostas Y1, Y2 e Y3 foram definidas como " & TextVar1 & ", " & TextVar2 & " e " & TextVar3 & "!")
    End If

End Sub

