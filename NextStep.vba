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
        
'
' Delete Graduate Students Entire Row
'
    
    For i = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
        If Cells(i, 11) = "graduate student" Then Rows(i).Delete
    Next
    
    For i = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
        If Cells(i, 11) = "Graduate Student" Then Rows(i).Delete
    Next

'
' Change 'United States' to 'US'
'

    Columns("H:H").Select
    Selection.Replace What:="United States", Replacement:="US", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                
'
' Delete all -999 fields in College CEEB Code Field
'

    Columns("R:R").Select
    Selection.Replace What:="-999", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
'
' Delete all -999 fields in High School CEEB Code Field
'

    Columns("M:M").Select
    Selection.Replace What:="-999", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'
' Delete all 9999 fields in High School Grad Date Column
'
    
    Columns("O:O").Select
    Selection.Replace What:="9999", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
'
' Delete all 99/99/9999 fields in BirthDate
'
    Columns("L:L").Select
    Selection.Replace What:="99/99/9999", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'
' Fix Birthday column to follow proper formatting
'
    Columns("L:L").EntireColumn.AutoFit
    Columns("L:L").Select
    Selection.NumberFormat = "yyyy/mm/dd"
    
'
' Need to Select only first interest
'

    Columns("S:S").Replace What:=",*", Replacement:=""
    Columns("S:S").Replace What:="/*", Replacement:=""
    Columns("S:S").Replace What:="&*", Replacement:=""
        
'
' Fix Major Column to match with Banner Code
'

    Columns("S:S").Replace What:="Advertising", Replacement:="ARGD_BFA"
    Columns("S:S").Replace What:="Business", Replacement:="BUAD_BBA"
    Columns("S:S").Replace What:="Criminal Justice", Replacement:="CRIJ_BS"
    Columns("S:S").Replace What:="Accounting", Replacement:="ACCT_BBA"
    Columns("S:S").Replace What:="Accounting and Business", Replacement:="ACCT_BBA"
    Columns("S:S").Replace What:="Accounting and Computer Science", Replacement:="ACCT_BBA"
    Columns("S:S").Replace What:="Accounting and Finance", Replacement:="ACCT_BBA"
    Columns("S:S").Replace What:="Accounting and Related Services", Replacement:="ACCT_BBA"
    Columns("S:S").Replace What:="Accounting Technology", Replacement:="ACCT_BBA"
    Columns("S:S").Replace What:="Acoustics", Replacement:="AASC_BAAS_SC"
    Columns("S:S").Replace What:="Acting", Replacement:="THEA_BFA_FM"
    Columns("S:S").Replace What:="Actuarial Science", Replacement:="AASC_BAAS_SC"
    Columns("S:S").Replace What:="Acupuncture and Oriental Medicine", Replacement:="BIOL_BS_NURS"
    Columns("S:S").Replace What:="Administration of Special Education", Replacement:="AASC_BAAS_SC"
    Columns("S:S").Replace What:="Administrative Assistant and Secretarial Science", Replacement:="BUAD_BBA"
    Columns("S:S").Replace What:="Adult and Continuing Education Administration", Replacement:="AASC_BAAS_SC"
    Columns("S:S").Replace What:="Adult and continuing Education and Teaching", Replacement:="INST_BS"
    Columns("S:S").Replace What:="Adult Development and Aging", Replacement:="HLTH_BS"
    Columns("S:S").Replace What:="Adult Health Nurse", Replacement:="BIOL_BS_NURS"
    Columns("S:S").Replace What:="Adult Literacy Tutor", Replacement:="INST_BS"
    
    
    
        
    
End Sub
