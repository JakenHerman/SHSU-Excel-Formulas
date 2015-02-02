Sub NextStep()
'
' NextStep Macro
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
    
    Columns("V:V").Select
    Selection.Replace What:="Yes", Replacement:="Y", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Columns("V:V").Select
    Selection.Replace What:="No", Replacement:="N", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Columns("U:U").Select
    Selection.Replace What:="Black/African American", Replacement:="2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Columns("U:U").Select
    Selection.Replace What:="Asian or Pacific Islander", Replacement:="5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Columns("U:U").Select
    Selection.Replace What:="White/Caucasian", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Columns("U:U").Select
    Selection.Replace What:="Spanish/Hispanic/Latino", Replacement:="4", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Columns("U:U").Select
    Selection.Replace What:="American Indian", Replacement:="3", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                
End Sub
