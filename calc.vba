Private Sub Worksheet_Change(ByVal Target As Range)

Application.EnableEvents = True
ThisWorkbook.RefreshAll

If Not Intersect(Target, Range("C3:C7")) Is Nothing Then
Call sbClearCells
Call Go
End If

End Sub
