Sub NextStep()

'
'Delete unneccesary columns and added needed ones
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
    
'
' Replace Yes and No Values to Y and N
'
    
    Columns("V:V").Select
    Selection.Replace What:="Yes", Replacement:="Y", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Columns("V:V").Select
    Selection.Replace What:="No", Replacement:="N", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
'
' Replace Ethinicity Text Values with Numeric Values
'

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
        
    Columns("U:U").Select
    Selection.Replace What:="Other", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Columns("U:U").Select
    Selection.Replace What:="International", Replacement:="6", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Columns("U:U").Select
    Selection.Replace What:="", Replacement:="7", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'
' Change Student Type to Either F or T
'
    Columns("K:K").Select
    Selection.Replace What:="High School Senior", Replacement:="F", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Columns("K:K").Select
    Selection.Replace What:="High School Junior", Replacement:="F", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Columns("K:K").Select
    Selection.Replace What:="High School Sophomore", Replacement:="F", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Columns("K:K").Select
    Selection.Replace What:="High School Freshman", Replacement:="F", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Columns("K:K").Select
    Selection.Replace What:="High School Senior", Replacement:="F", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Columns("K:K").Select
    Selection.Replace What:="College Senior", Replacement:="T", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Columns("K:K").Select
    Selection.Replace What:="College Junior", Replacement:="T", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Columns("K:K").Select
    Selection.Replace What:="College Sophomore", Replacement:="T", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Columns("K:K").Select
    Selection.Replace What:="College Freshman", Replacement:="T", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Columns("K:K").Select
    Selection.Replace What:="Adult Learner", Replacement:="F", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Columns("K:K").Select
    Selection.Replace What:="Transfer Student", Replacement:="T", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
                
End Sub
