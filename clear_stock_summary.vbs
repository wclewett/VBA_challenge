Sub clear_contents():
Dim ws_count As Integer
Dim ws As Integer

' Set WS_Count equal to the number of worksheets in the active workbook.
ws_count = ActiveWorkbook.Worksheets.Count

' Loop through worksheets
For ws = 1 To ws_count

worksheet_name = ActiveWorkbook.Worksheets(ws).Name
Worksheets(worksheet_name).Activate

Columns("I:P").EntireColumn.Delete

Next ws

End Sub