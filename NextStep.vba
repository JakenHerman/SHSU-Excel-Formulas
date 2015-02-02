Sub NextStep()
'
' Macro2 Macro
'

'
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Student Type"
    
    Columns("P:P").Select
    Selection.Delete Shift:=xlToLeft
    
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Entry Term"
    
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Entry Year"
    
    Columns("S:T").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Major 1"
    
    Columns("T:T").Select
    Selection.Delete Shift:=xlToLeft

    Columns("Z:AB").Select
    Selection.Delete Shift:=xlToLeft
    
End Sub
