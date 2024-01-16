# VBA-COMMANDS
My most frequent used commands in VBA

## GET FIRST & LAST USED ROW IN A WORKSHEET
```
R1 =  Worksheets("Datos").Cells(Rows.Count, 1).End(xlUp).Row 'Fin serie

R2 = Worksheets("Datos").Cells(R1, "Q").End(xlUp).Row   'Inicio serie



```
## GET COLUMN NUMBER GIVEN THE NAME OF THE COLUMN
```
Sub ExampleUsage()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change the sheet name as needed
    
    Dim columnName As String
    columnName = "ColumnName" ' Change the column name as needed
    
    Dim columnNumber As Long
    columnNumber = GetColumnNumber(ws, columnName)
    
    If columnNumber > 0 Then
        MsgBox "Column '" & columnName & "' is in column number " & columnNumber & "."
    Else
        MsgBox "Column '" & columnName & "' not found in the specified worksheet."
    End If
End Sub
```
