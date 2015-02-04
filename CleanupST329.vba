Sub CleanupST329()

'
' CleanupST329 Macro
'
    Rows("1:3").Select
    Range("V3").Activate
    Selection.Delete Shift:=xlUp
    Columns("O:T").Select
    Range("T1").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("C:J").Select
    Range("J1").Activate
    Selection.Delete Shift:=xlToLeft
End Sub
