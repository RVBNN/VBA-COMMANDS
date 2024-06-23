# VBA-COMMANDS
My most frequent used commands in VBA

## TURN OFF FLASHES AND AUTOMATIC CALCULATIONS
```
    Application.ScreenUpdating = False 'Turn off Flashes
    Application.Calculation = xlCalculationManual 'Turn off automatic calculations

    ' Turn them back on
    'Application.ScreenUpdating = True ' Flashes
    'Application.Calculation = xlCalculationAutomatic 'Automatic calculations
```


## GET FIRST & LAST USED ROW IN A WORKSHEET
```
Dim datos_ws as Worksheets
Set datos_ws = Worksheets("Datos")

R1 =  datos_ws.Cells(Rows.Count, 1).End(xlUp).Row 'Fin serie

R2 = datos_ws.Cells(R1, "Q").End(xlUp).Row   'Inicio serie

lastUsedColumn = datos_ws.Cells(1, datos_ws.Columns.Count).End(xlToLeft).Column

```

## ONLY VISIBLE ROWS
```
Selection.SpecialCells(xlCellTypeVisible).Select
```

## GET TABLE RANGE
```
Sub table_range()
    Dim det_ws As Worksheet
    Dim det_range As Range
    Set det_ws = ActiveWorkbook.Worksheets("DETALLE")

' Obtenemos el rango en el cual hay valores
    lastUsedColumn_det = det_ws.Cells(1, det_ws.Columns.Count).End(xlToLeft).Column

    lastRow_det = det_ws.Cells(Rows.Count, 1).End(xlUp).Row
    lastRow_address_det = det_ws.Cells(lastRow_det, lastUsedColumn_det).Address

    firstRow_det = det_ws.Cells(lastRow_det, 1).End(xlUp).Row 'EN OCASIONES HAY QUE SUMARLE UNO PORQUE ES EL ENCABEZADO
    firstRow_address_det = det_ws.Cells(firstRow_det, 1).Address

    Set det_range = det_ws.Range(firstRow_address_det, lastRow_address_det)
End Sub
```

## INSERT FORMULAS IN CELLS
```    
    Range("H2:H" & ultimo).FormulaLocal = "=I2" 'Lo que sea que tenga en la columna I se lo pone a la columna H

    ' If you're going to use a dynamic range, then you MUST use the Set syntax -------------------------------
    ' Set myRange = ws.Range(Inicio, Fin)
    ' myRange.FormulaLocal = "=I2"
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

## WORD DOCUMENT MANIPULATION WITH EXCEL MACROS (CREATION, ALIGNMENT, STYLE)
```
    Sub CreateWordDoc()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wdRange As Object

    ' Create a new instance of Word application
    Set wdApp = CreateObject("Word.Application")
    
    ' Make Word visible
    wdApp.Visible = True
    
    ' Add a new document
    Set wdDoc = wdApp.Documents.Add
    
    ' Insert and format text
    With wdDoc
        ' Centered text
        Set wdRange = .Range
        wdRange.Text = "This is centered text."
        wdRange.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
        wdRange.InsertParagraphAfter
        
        ' Right aligned text
        Set wdRange = .Paragraphs(.Paragraphs.Count).Range
        wdRange.Text = "This is right aligned text."
        wdRange.ParagraphFormat.Alignment = 3 ' wdAlignParagraphRight
        wdRange.InsertParagraphAfter
        
        ' Justified text
        Set wdRange = .Paragraphs(.Paragraphs.Count).Range
        wdRange.Text = "This is justified text. " & _
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " & _
            "Curabitur pretium tincidunt lacus. Nulla gravida orci a odio."
        wdRange.ParagraphFormat.Alignment = 3 ' wdAlignParagraphJustify
        wdRange.InsertParagraphAfter
        
        ' Left aligned text
        Set wdRange = .Paragraphs(.Paragraphs.Count).Range
        wdRange.Text = "This is left aligned text."
        wdRange.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
        wdRange.InsertParagraphAfter
        
        ' Bold text
        Set wdRange = .Paragraphs(.Paragraphs.Count).Range
        wdRange.Text = "This is bold text."
        wdRange.Font.Bold = True
        wdRange.InsertParagraphAfter
        
        ' Italic text
        Set wdRange = .Paragraphs(.Paragraphs.Count).Range
        wdRange.Text = "This is italic text."
        wdRange.Font.Italic = True
        wdRange.InsertParagraphAfter
        
        ' Hyperlink
        Set wdRange = .Paragraphs(.Paragraphs.Count).Range
        wdRange.Text = "This is a hyperlink."
        .Hyperlinks.Add Anchor:=wdRange, Address:="http://www.example.com", TextToDisplay:="This is a hyperlink"
        wdRange.InsertParagraphAfter
    End With
    
    ' Clean up
    Set wdRange = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
```




