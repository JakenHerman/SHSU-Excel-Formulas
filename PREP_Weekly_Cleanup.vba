Sub Cleanup()
'
' PREP Weekly Cleanup Macro
'

'
    ActiveSheet.Shapes.Range(Array("Big_S.jpeg")).Select
    Selection.Delete
    Rows("9:9").RowHeight = 12
    Rows("1:10").Select
    Range("A10").Activate
    Selection.Delete Shift:=xlUp
    Columns("C:M").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
End Sub
