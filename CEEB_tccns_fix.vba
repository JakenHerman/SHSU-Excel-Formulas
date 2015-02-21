Sub CEEB_tccns_fix()
'
' CEEB_tccns_fix Macro
'

'
   
For i = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
    If Cells(i, 5) = "----" Then Rows(i).Delete
Next



End Sub
